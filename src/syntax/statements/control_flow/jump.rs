//! Jump statement parsing for VB6 (`GoTo`, `GoSub`, `Return`, `Labels`).

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a `GoSub` statement.
    ///
    /// VB6 `GoSub` statement syntax:
    /// - `GoSub` label
    ///
    /// Branches to and returns from a subroutine within a procedure.
    ///
    /// The `GoSub`...`Return` statement syntax has these parts:
    ///
    /// | Part   | Description |
    /// |--------|-------------|
    /// | label  | Required. A line label or line number. |
    ///
    /// Remarks:
    /// - You can use `GoSub` and `Return` anywhere in a procedure, but `GoSub` and the corresponding `Return` statement must be in the same procedure.
    /// - A subroutine can contain more than one `Return` statement, but the first one encountered causes the flow of execution to branch back to the statement immediately following the most recently executed `GoSub` statement.
    /// - You can't enter or exit `Sub` procedures with `GoSub`...`Return`.
    /// - Using `GoSub` and `Return` is considered obsolete. Modern VB6 code should use `Sub` or `Function` procedures instead.
    ///
    /// Examples:
    /// ```vb
    /// Sub Test()
    ///     GoSub ErrorHandler
    ///     Exit Sub
    /// ErrorHandler:
    ///     MsgBox "Error"
    ///     Return
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/gosubreturn-statement)
    pub(crate) fn parse_gosub_statement(&mut self) {
        // if we are now parsing a gosub statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::GoSubStatement.to_raw());
        self.consume_whitespace();

        // Consume "GoSub" keyword
        self.consume_token();

        // Consume everything until newline (the label name)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // GoSubStatement
    }

    /// Parse a Return statement.
    ///
    /// VB6 Return statement syntax:
    /// - Return
    ///
    /// Returns from a subroutine within a procedure.
    ///
    /// Remarks:
    /// - `Return` must be used with `GoSub` to return to the statement following the `GoSub` call.
    /// - You can use `GoSub` and `Return` anywhere in a procedure, but `GoSub` and the corresponding `Return` statement must be in the same procedure.
    /// - A subroutine can contain more than one `Return` statement, but the first one encountered causes the flow of execution to branch back to the statement immediately following the most recently executed `GoSub` statement.
    /// - Using `GoSub` and `Return` is considered obsolete. Modern VB6 code should use `Sub` or `Function` procedures instead.
    ///
    /// Examples:
    /// ```vb
    /// Sub Test()
    ///     GoSub Cleanup
    ///     Exit Sub
    /// Cleanup:
    ///     Set obj = Nothing
    ///     Return
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/gosubreturn-statement)
    pub(crate) fn parse_return_statement(&mut self) {
        // if we are now parsing a return statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ReturnStatement.to_raw());
        self.consume_whitespace();

        // Consume "Return" keyword
        self.consume_token();

        // Consume the newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ReturnStatement
    }

    /// Parse a `GoTo` statement.
    ///
    /// Syntax:
    ///   `GoTo` label
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(crate) fn parse_goto_statement(&mut self) {
        // if we are now parsing a `GoTo` statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::GotoStatement.to_raw());
        self.consume_whitespace();

        // Consume "`GoTo`" keyword
        self.consume_token();

        // Consume everything until newline (the label name)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // GotoStatement
    }

    /// Parse a label statement.
    ///
    /// VB6 label syntax:
    /// - `LabelName:`
    ///
    /// `Labels` are used as targets for `GoTo` and `GoSub` statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(crate) fn parse_label_statement(&mut self) {
        // if we are now parsing a label statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::LabelStatement.to_raw());
        self.consume_whitespace();

        // Consume the label identifier
        self.consume_token();

        // Consume optional whitespace
        self.consume_whitespace();

        // Consume the colon
        if self.at_token(Token::ColonOperator) {
            self.consume_token();
        }

        // Consume the newline if present
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LabelStatement
    }

    /// Check if the current position is at a label.
    /// A label is an identifier or number followed by a colon.
    pub(crate) fn is_at_label(&self) -> bool {
        let next_token_is_colon = matches!(self.peek_next_token(), Some(Token::ColonOperator));

        if !next_token_is_colon {
            return false;
        }

        // If we are not parsing the header, then some keywords are valid identifiers (like "Begin")
        // TODO: Consider adding a list of keywords that can be used as labels.
        // TODO: Also consider modifying tokenizer to recognize when inside header to more easily identify Identifiers vs header only keywords.
        if !self.parsing_header && matches!(self.current_token(), Some(Token::BeginKeyword)) {
            return true;
        }

        self.is_identifier() || self.is_number()
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    // GoTo statement tests

    #[test]
    fn goto_statement_simple() {
        let source = r#"
Sub Test()
    GoTo ErrorHandler
    Debug.Print "Normal code"
ErrorHandler:
    Debug.Print "Error handling"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_with_line_number() {
        let source = r#"
Sub Test()
    GoTo 100
    Debug.Print "code"
100:
    Debug.Print "target"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_in_if() {
        let source = r#"
Sub Test()
    If x > 10 Then
        GoTo LargeValue
    End If
    Debug.Print "small"
LargeValue:
    Debug.Print "large"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_multiple() {
        let source = r#"
Sub Test()
    GoTo Label1
    GoTo Label2
    GoTo Label3
Label1:
    Debug.Print "one"
Label2:
    Debug.Print "two"
Label3:
    Debug.Print "three"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_error_handling() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Some code that might error
    Debug.Print "normal"
    Exit Sub
ErrorHandler:
    MsgBox "Error occurred"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_forward_jump() {
        let source = r"
Sub Test()
    x = 1
    GoTo SkipCode
    x = 2
    x = 3
SkipCode:
    x = 4
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_backward_jump() {
        let source = r"
Sub Test()
StartLoop:
    counter = counter + 1
    If counter < 10 Then
        GoTo StartLoop
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_in_select_case() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1
            GoTo Handler1
        Case 2
            GoTo Handler2
    End Select
Handler1:
    Debug.Print "one"
Handler2:
    Debug.Print "two"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 5 Then
            GoTo ExitLoop
        End If
        Debug.Print i
    Next i
ExitLoop:
    Debug.Print "done"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_module_level() {
        let source = r#"
Public Sub TestGoto()
    GoTo Finish
    Debug.Print "skipped"
Finish:
    Debug.Print "done"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_with_underscore() {
        let source = r#"
Sub Test()
    GoTo Error_Handler
Error_Handler:
    Debug.Print "error"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    GoTo MyLabel
MyLabel:
    x = 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_same_line_as_then() {
        let source = r"
Sub Test()
    If condition Then
        GoTo Handler
    End If
Handler:
    result = True
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn goto_statement_exit_cleanup() {
        let source = r"
Sub Test()
    On Error GoTo Cleanup
    ' do work
    Exit Sub
Cleanup:
    ' cleanup code
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    // Label statement tests

    #[test]
    fn label_simple() {
        let source = r"
Sub Test()
    MyLabel:
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_with_goto() {
        let source = r#"
Sub Test()
    GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    MsgBox "Error"
End Sub
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_with_underscore() {
        let source = r#"
Sub Test()
Error_Handler:
    MsgBox "Error"
End Sub
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_at_module_level() {
        let source = r"
Sub Test()
StartHere:
    x = 1
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_multiple() {
        let source = r"
Sub Test()
Start:
    x = 1
Middle:
    y = 2
End_Label:
    z = 3
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_with_space_after_colon() {
        let source = r"
Sub Test()
MyLabel: x = 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_error_handler_pattern() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    ' Some code
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_with_numbers() {
        let source = r"
Sub Test()
Label123:
    x = 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_cleanup_pattern() {
        let source = r"
Sub Test()
    GoTo Cleanup
Cleanup:
    Set obj = Nothing
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_preserves_whitespace() {
        let source = "MyLabel:";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_in_function() {
        let source = r"
Function Calculate() As Integer
Start:
    Calculate = 42
End Function
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_mixed_case() {
        let source = r#"
Sub Test()
MyErrorHandler:
    MsgBox "Error"
End Sub
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    // GoSub statement tests

    #[test]
    fn gosub_simple() {
        let source = r#"
Sub Test()
    GoSub ErrorHandler
    Exit Sub
ErrorHandler:
    MsgBox "Error"
    Return
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_with_line_number() {
        let source = r#"
Sub Test()
    GoSub 100
    Exit Sub
100:
    Debug.Print "subroutine"
    Return
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_in_if_statement() {
        let source = r"
Sub Test()
    If x > 0 Then
        GoSub ProcessPositive
    End If
    Exit Sub
ProcessPositive:
    y = y + 1
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_multiple() {
        let source = r#"
Sub Test()
    GoSub Sub1
    GoSub Sub2
    Exit Sub
Sub1:
    Debug.Print "one"
    Return
Sub2:
    Debug.Print "two"
    Return
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_nested_calls() {
        let source = r#"
Sub Test()
    GoSub Level1
    Exit Sub
Level1:
    GoSub Level2
    Return
Level2:
    Debug.Print "deep"
    Return
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_preserves_whitespace() {
        let source = r"
Sub Test()
    GoSub   MyLabel
MyLabel:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_with_comment() {
        let source = r"
Sub Test()
    GoSub Cleanup ' Call cleanup routine
Cleanup:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_inline_if() {
        let source = r"
Sub Test()
    If needsInit Then GoSub Initialize
    Exit Sub
Initialize:
    x = 0
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        GoSub ProcessItem
    Next i
    Exit Sub
ProcessItem:
    Debug.Print i
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_in_select_case() {
        let source = r"
Sub Test()
    Select Case x
        Case 1
            GoSub Handler1
        Case 2
            GoSub Handler2
    End Select
    Exit Sub
Handler1:
    Return
Handler2:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_with_underscore_label() {
        let source = r"
Sub Test()
    GoSub Error_Handler
    Exit Sub
Error_Handler:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_error_handling_pattern() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorExit
    GoSub DoWork
    Exit Sub
DoWork:
    ' work code
    Return
ErrorExit:
    MsgBox "Error"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    // Return statement tests

    #[test]
    fn return_simple() {
        let source = r"
Sub Test()
    GoSub SubRoutine
    Exit Sub
SubRoutine:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_multiple() {
        let source = r"
Sub Test()
    GoSub Process
    Exit Sub
Process:
    If x > 0 Then
        Return
    End If
    y = 1
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_in_if_statement() {
        let source = r"
Sub Test()
    GoSub Check
    Exit Sub
Check:
    If x = 0 Then
        Return
    End If
    x = x + 1
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_inline_if() {
        let source = r"
Sub Test()
    GoSub Validate
    Exit Sub
Validate:
    If invalid Then Return
    DoSomething
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_with_comment() {
        let source = r"
Sub Test()
    GoSub Cleanup
    Exit Sub
Cleanup:
    Set obj = Nothing
    Return ' Exit subroutine
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_preserves_whitespace() {
        let source = r"
Sub Test()
    GoSub Sub1
    Exit Sub
Sub1:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_in_select_case() {
        let source = r"
Sub Test()
    GoSub Process
    Exit Sub
Process:
    Select Case x
        Case 1
            Return
        Case 2
            y = 2
    End Select
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_in_loop() {
        let source = r"
Sub Test()
    GoSub FindValue
    Exit Sub
FindValue:
    For i = 1 To 10
        If arr(i) = target Then Return
    Next i
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_return_complete_example() {
        let source = r"
Sub Main()
    x = 10
    GoSub DoubleValue
    Debug.Print x
    Exit Sub
DoubleValue:
    x = x * 2
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_return_nested_example() {
        let source = r#"
Sub Test()
    GoSub Outer
    Exit Sub
Outer:
    GoSub Inner
    Return
Inner:
    Debug.Print "nested"
    Return
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn return_at_module_level() {
        let source = r"
Public Sub TestReturn()
    GoSub Handler
    Exit Sub
Handler:
    Return
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn gosub_return_error_pattern() {
        let source = r"
Sub Test()
    On Error GoTo ErrHandler
    GoSub ProcessData
    Exit Sub
ProcessData:
    ' process
    Return
ErrHandler:
    MsgBox Err.Description
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/jump");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
