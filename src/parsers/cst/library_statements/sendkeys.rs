//! # `SendKeys` Statement
//!
//! Sends one or more keystrokes to the active window as if typed at the keyboard.
//!
//! ## Syntax
//!
//! ```vb
//! SendKeys string [, wait]
//! ```
//!
//! ## Parts
//!
//! - **string**: Required. String expression specifying the keystrokes to send.
//! - **wait**: Optional. Boolean value specifying the wait mode. If True, Visual Basic waits for the keystrokes to be processed before returning control to the calling procedure. If False (default), control returns immediately after the keys are sent.
//!
//! ## Remarks
//!
//! - **Active Window**: `SendKeys` sends keystrokes to the currently active window. Your application must activate the target window before using `SendKeys`.
//! - **Keystroke Representation**: Each key is represented by one or more characters. To specify a single keyboard character, use the character itself (e.g., "A" sends the letter A).
//! - **Multiple Characters**: To send a string of characters, concatenate them (e.g., "Hello" sends H, e, l, l, o in sequence).
//! - **Special Keys**: Some keys have special representations enclosed in braces (e.g., {ENTER}, {TAB}, {ESC}).
//! - **Wait Parameter**: Setting wait to True ensures that keystrokes are processed before your code continues. This is useful when you need to wait for an application to respond.
//! - **Focus Issues**: If the target application doesn't have focus when `SendKeys` executes, the keystrokes may be sent to the wrong application.
//! - **`AppActivate`**: Use `AppActivate` to activate the target window before calling `SendKeys`.
//!
//! ## Special Key Codes
//!
//! | Key | Code |
//! |-----|------|
//! | BACKSPACE | {BACKSPACE} or {BS} or {BKSP} |
//! | BREAK | {BREAK} |
//! | CAPS LOCK | {CAPSLOCK} |
//! | DELETE | {DELETE} or {DEL} |
//! | DOWN ARROW | {DOWN} |
//! | END | {END} |
//! | ENTER | {ENTER} or ~ |
//! | ESC | {ESC} or {ESCAPE} |
//! | HELP | {HELP} |
//! | HOME | {HOME} |
//! | INSERT | {INSERT} or {INS} |
//! | LEFT ARROW | {LEFT} |
//! | NUM LOCK | {NUMLOCK} |
//! | PAGE DOWN | {PGDN} |
//! | PAGE UP | {PGUP} |
//! | PRINT SCREEN | {PRTSC} |
//! | RIGHT ARROW | {RIGHT} |
//! | SCROLL LOCK | {SCROLLLOCK} |
//! | TAB | {TAB} |
//! | UP ARROW | {UP} |
//! | F1-F16 | {F1} through {F16} |
//!
//! ## Modifier Keys
//!
//! | Key | Code |
//! |-----|------|
//! | SHIFT | + (plus sign) |
//! | CTRL | ^ (caret) |
//! | ALT | % (percent sign) |
//!
//! To specify modifier keys with regular keys, enclose the regular keys in parentheses:
//! - `"+{F1}"` sends SHIFT+F1
//! - `"^(ec)"` sends CTRL+E followed by CTRL+C
//! - `"%(FA)"` sends ALT+F followed by ALT+A
//!
//! ## Repeating Keys
//!
//! To repeat a key, use the format `{key number}`:
//! - `"{RIGHT 10}"` sends RIGHT arrow 10 times
//! - `"{TAB 5}"` sends TAB 5 times
//!
//! ## Examples
//!
//! ### Send Simple Text
//!
//! ```vb
//! SendKeys "Hello World"
//! ```
//!
//! ### Send Text with Enter Key
//!
//! ```vb
//! SendKeys "Username{TAB}Password{ENTER}"
//! ```
//!
//! ### Activate Window and Send Keys
//!
//! ```vb
//! AppActivate "Notepad"
//! SendKeys "Hello from VB6{ENTER}", True
//! ```
//!
//! ### Send Alt+F4 to Close Window
//!
//! ```vb
//! SendKeys "%{F4}"  ' ALT+F4
//! ```
//!
//! ### Send Ctrl+C to Copy
//!
//! ```vb
//! SendKeys "^c"  ' CTRL+C
//! ```
//!
//! ### Send Multiple Keys with Wait
//!
//! ```vb
//! SendKeys "{DOWN}{DOWN}{ENTER}", True
//! ```
//!
//! ### Fill Form Fields
//!
//! ```vb
//! AppActivate "Data Entry Form"
//! SendKeys "John Doe{TAB}123 Main St{TAB}555-1234{ENTER}", True
//! ```
//!
//! ### Send Function Keys
//!
//! ```vb
//! SendKeys "{F1}"    ' Help key
//! SendKeys "{F5}"    ' Refresh
//! SendKeys "+{F10}"  ' SHIFT+F10 (context menu)
//! ```
//!
//! ### Repeat Keys
//!
//! ```vb
//! SendKeys "{RIGHT 5}"    ' Move right 5 times
//! SendKeys "{DOWN 10}"    ' Move down 10 times
//! SendKeys "{BACKSPACE 3}" ' Delete 3 characters
//! ```
//!
//! ### Send Key Combinations
//!
//! ```vb
//! SendKeys "^a"       ' CTRL+A (Select All)
//! SendKeys "^c"       ' CTRL+C (Copy)
//! SendKeys "^v"       ' CTRL+V (Paste)
//! SendKeys "^s"       ' CTRL+S (Save)
//! ```
//!
//! ### Navigate Menus
//!
//! ```vb
//! AppActivate "Microsoft Word"
//! SendKeys "%f", True  ' ALT+F (File menu)
//! SendKeys "s", True   ' S (Save)
//! ```
//!
//! ### Send Special Characters
//!
//! ```vb
//! SendKeys "Test {+} Addition"  ' Sends: Test + Addition
//! SendKeys "Test {^} Power"     ' Sends: Test ^ Power
//! SendKeys "Test {% } Percent"  ' Sends: Test % Percent
//! ```
//!
//! ## Important Notes
//!
//! - **Timing**: `SendKeys` is not always reliable for complex automation. Consider using API calls or UI automation libraries for critical tasks.
//! - **Focus Management**: Always ensure the target window has focus before sending keys.
//! - **Wait Parameter**: Use True for the wait parameter when you need synchronous operation.
//! - **Case Sensitivity**: To send uppercase letters, use the SHIFT modifier: `"+abc"` sends uppercase ABC.
//! - **Reserved Characters**: To send +, ^, %, ~, or {}, enclose them in braces: `{+}`, `{^}`, `{%}`, `{~}`, `{{}`, `{}}`.
//! - **Limitations**: `SendKeys` doesn't work with applications that directly process keyboard input at a low level.
//! - **Error Handling**: If the target application is busy or unresponsive, `SendKeys` may fail silently or send keys to the wrong window.
//!
//! ## Common Errors
//!
//! - **Error 5**: Invalid procedure call - occurs if string contains invalid key codes
//! - Keys sent to wrong application if focus isn't properly managed
//! - Timing issues when wait is False and subsequent code depends on keystrokes being processed
//!
//! ## Best Practices
//!
//! - Always use `AppActivate` before `SendKeys` to ensure the correct window receives the keystrokes
//! - Use the wait parameter (True) when the next operation depends on the keystrokes being processed
//! - Add delays (`DoEvents` or Sleep) between `SendKeys` calls for complex sequences
//! - Test thoroughly as `SendKeys` behavior can vary across different applications and Windows versions
//! - Consider alternatives like Windows API or UI Automation for production applications
//!
//! ## See Also
//!
//! - `AppActivate` statement (activate an application window)
//! - `DoEvents` function (yield execution to allow events to be processed)
//! - `Shell` function (run executable programs)
//!
//! ## References
//!
//! - [SendKeys Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sendkeys-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `SendKeys` statement.
    pub(crate) fn parse_sendkeys_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::SendKeysStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn sendkeys_simple() {
        let source = r#"
Sub Test()
    SendKeys "Hello World"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("SendKeysKeyword"));
    }

    #[test]
    fn sendkeys_at_module_level() {
        let source = "SendKeys \"Test\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_wait_true() {
        let source = r#"
Sub Test()
    SendKeys "Text", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("True"));
    }

    #[test]
    fn sendkeys_with_wait_false() {
        let source = r#"
Sub Test()
    SendKeys "Text", False
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("False"));
    }

    #[test]
    fn sendkeys_with_special_keys() {
        let source = r#"
Sub Test()
    SendKeys "{ENTER}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_tab() {
        let source = r#"
Sub Test()
    SendKeys "Username{TAB}Password{ENTER}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_variable() {
        let source = r"
Sub Test()
    SendKeys keyString
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("keyString"));
    }

    #[test]
    fn sendkeys_with_concatenation() {
        let source = r#"
Sub Test()
    SendKeys "Hello " & userName & "{ENTER}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("userName"));
    }

    #[test]
    fn sendkeys_alt_f4() {
        let source = r#"
Sub Test()
    SendKeys "%{F4}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_ctrl_c() {
        let source = r#"
Sub Test()
    SendKeys "^c"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_appactivate() {
        let source = r#"
Sub Test()
    AppActivate "Notepad"
    SendKeys "Hello", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn sendkeys_inside_if_statement() {
        let source = r#"
If needKeys Then
    SendKeys "{F5}"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_inside_loop() {
        let source = r#"
For i = 1 To 10
    SendKeys "{DOWN}"
Next i
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_comment() {
        let source = r#"
Sub Test()
    SendKeys "{ENTER}", True ' Press Enter
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("' Press Enter"));
    }

    #[test]
    fn sendkeys_preserves_whitespace() {
        let source = "SendKeys   \"Text\"  ,   True\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_function_keys() {
        let source = r#"
Sub Test()
    SendKeys "{F1}"
    SendKeys "{F5}"
    SendKeys "+{F10}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_repeat_keys() {
        let source = r#"
Sub Test()
    SendKeys "{RIGHT 5}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_in_select_case() {
        let source = r#"
Select Case action
    Case 1
        SendKeys "{F1}"
    Case 2
        SendKeys "{F5}"
End Select
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_multiple_on_same_line() {
        let source = "SendKeys \"A\": SendKeys \"B\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_wait_variable() {
        let source = r#"
Sub Test()
    SendKeys "Text", waitFlag
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("waitFlag"));
    }

    #[test]
    fn sendkeys_in_with_block() {
        let source = r"
With automation
    SendKeys .keySequence
End With
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
        assert!(debug.contains("keySequence"));
    }

    #[test]
    fn sendkeys_in_sub() {
        let source = r#"
Sub SendEnter()
    SendKeys "{ENTER}", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_in_function() {
        let source = r#"
Function AutomateInput() As Boolean
    SendKeys "Data{ENTER}", True
    AutomateInput = True
End Function
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_menu_navigation() {
        let source = r#"
Sub Test()
    SendKeys "%f", True
    SendKeys "s", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_in_class_module() {
        let source = r"
Private keyData As String

Public Sub SendData()
    SendKeys keyData, True
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_with_line_continuation() {
        let source = r#"
Sub Test()
    SendKeys _
        "Long text here", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_form_automation() {
        let source = r#"
Sub Test()
    SendKeys "John Doe{TAB}123 Main St{TAB}555-1234{ENTER}", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_arrows() {
        let source = r#"
Sub Test()
    SendKeys "{DOWN}{DOWN}{ENTER}"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_error_handling() {
        let source = r#"
On Error Resume Next
SendKeys keys, True
If Err.Number <> 0 Then
    MsgBox "SendKeys failed"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }

    #[test]
    fn sendkeys_empty_string() {
        let source = r#"
Sub Test()
    SendKeys ""
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SendKeysStatement"));
    }
}
