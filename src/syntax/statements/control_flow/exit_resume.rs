//! Exit and Resume statement parsing for VB6.

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a Resume statement.
    ///
    /// VB6 Resume statement syntax:
    /// - `Resume`
    /// - `Resume Next`
    /// - `Resume Label`
    ///
    /// Resumes execution after an error-handling routine is finished.
    ///
    /// # Syntax
    ///
    /// The `Resume` statement has these forms:
    ///
    /// | Form | Description |
    /// |------|-------------|
    /// | `Resume` | If the error occurred in the same procedure as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the statement that last called out of the procedure containing the error-handling routine. |
    /// | `Resume Next` | If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or On Error Resume Next statement). |
    /// | `Resume Label` | Execution resumes at the line specified by the label argument. The label argument can be a line label or line number. |
    ///
    /// # Remarks
    ///
    /// - The `Resume` statement can be used only in an error-handling routine.
    /// - Using `Resume` without specifying a label causes execution to resume at the statement that caused the error.
    /// - `Resume Next` is useful when you want to continue execution despite an error.
    /// - `Resume Label` is useful when you want to continue execution at a specific location after handling an error.
    /// - If you use a `Resume` statement anywhere except in an error-handling routine, an error occurs.
    /// - `Resume` cannot be used in any procedure that contains an On Error `Resume Next` statement.
    ///
    /// # Examples
    ///
    /// ```vb
    /// Sub Test()
    ///     On Error GoTo ErrorHandler
    ///     ' Code that might cause error
    ///     x = 1 / 0
    ///     Exit Sub
    /// ErrorHandler:
    ///     MsgBox "Error occurred"
    ///     Resume Next
    /// End Sub
    /// ```
    ///
    /// ```vb
    /// Sub Test2()
    ///     On Error GoTo ErrorHandler
    ///     ' Code that might cause error
    ///     Exit Sub
    /// ErrorHandler:
    ///     If Err.Number = 11 Then
    ///         Resume
    ///     Else
    ///         Resume CleanUp
    ///     End If
    /// CleanUp:
    ///     ' Cleanup code
    /// End Sub
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft VBA Language Reference - Resume Statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement)
    pub(crate) fn parse_resume_statement(&mut self) {
        // if we are now parsing a resume statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ResumeStatement.to_raw());
        self.consume_whitespace();

        // Consume "Resume" keyword
        self.consume_token();

        // Consume everything until newline (Next keyword or label)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // ResumeStatement
    }

    /// Parse an Exit statement.
    ///
    /// VB6 Exit statement syntax:
    /// - Exit Do
    /// - Exit For
    /// - Exit Function
    /// - Exit Property
    /// - Exit Sub
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/exit-statement)
    pub(crate) fn parse_exit_statement(&mut self) {
        // if we are now parsing an exit statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ExitStatement.to_raw());
        self.consume_whitespace();

        // Consume "Exit" keyword
        self.consume_token();

        // Consume whitespace after Exit
        self.consume_whitespace();

        // Consume the exit type (Do, For, Function, Property, Sub)
        if self.at_token(Token::DoKeyword)
            || self.at_token(Token::ForKeyword)
            || self.at_token(Token::FunctionKeyword)
            || self.at_token(Token::PropertyKeyword)
            || self.at_token(Token::SubKeyword)
        {
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ExitStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn exit_do() {
        let source = r"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_for() {
        let source = r"
Sub Test()
    For i = 1 To 10
        If i = 5 Then Exit For
    Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_function() {
        let source = r"
Function Test() As Integer
    If x = 0 Then
        Exit Function
    End If
    Test = 42
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_sub() {
        let source = r#"
Sub Test()
    If x = 0 Then Exit Sub
    Debug.Print "x is not zero"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_property() {
        let source = r"
Property Set Callback(ByRef newObj As InterPress)
    Set mCallback = newObj
    Exit Property
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_exit_statements() {
        let source = r"
Sub Test()
    For i = 1 To 10
        If i = 3 Then Exit For
        If i = 7 Then Exit Sub
    Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_in_nested_loops() {
        let source = r"
Sub Test()
    Do While x < 100
        For i = 1 To 10
            If i = 5 Then Exit For
        Next
        If x > 50 Then Exit Do
        x = x + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn exit_preserves_whitespace() {
        let source = r"
Sub Test()
    Exit   Sub
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_exit_in_if_statement() {
        let source = r"
Function Test(x As Integer) As Integer
    If x = 0 Then Exit Function
    Test = x * 2
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    // Resume statement tests

    #[test]
    fn resume_simple() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_next() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    x = 1 / 0
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_label() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume CleanUp
CleanUp:
    MsgBox "Cleanup"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_line_number() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume 100
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_in_error_handler() {
        let source = r#"
Sub ProcessFile()
    On Error GoTo FileError
    Open "test.txt" For Input As #1
    Exit Sub
FileError:
    If Err.Number = 53 Then
        MsgBox "File not found"
        Resume Next
    Else
        Resume
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_comment() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume Next ' Continue after error
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_preserves_whitespace() {
        let source = "    Resume    Next    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_in_nested_error_handler() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Call SubProcedure
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then
        Resume Next
    Else
        Resume CleanUp
    End If
CleanUp:
    ' Cleanup code
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_at_module_level() {
        let source = "Resume Next\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_select_case() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 5
            Resume Next
        Case 11
            Resume
        Case Else
            Resume CleanUp
    End Select
CleanUp:
    ' Cleanup
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_in_loop() {
        let source = r"
Sub ProcessFiles()
    On Error GoTo ErrorHandler
    For i = 1 To 10
        ProcessFile i
    Next i
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_complex_error_handling() {
        let source = r#"
Function OpenDatabase() As Boolean
    On Error GoTo DBError
    ' Database opening code
    OpenDatabase = True
    Exit Function
DBError:
    Select Case Err.Number
        Case 3024
            MsgBox "Database locked"
            Resume Retry
        Case 3044
            MsgBox "Path not found"
            Resume Next
        Case Else
            MsgBox "Unknown error"
            Resume ExitPoint
    End Select
Retry:
    ' Retry logic
    Resume
ExitPoint:
    OpenDatabase = False
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_do_loop() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Do
        ' Code
    Loop
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_resume_statements() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then
        Resume
    End If
    If Err.Number = 11 Then
        Resume Next
    End If
    Resume CleanUp
CleanUp:
    ' Cleanup
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_on_error_resume_next() {
        let source = r#"
Sub Test()
    On Error Resume Next
    x = 1 / 0
    If Err.Number <> 0 Then
        MsgBox "Error occurred"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_in_class_module() {
        let source = r"
Private Sub Class_Initialize()
    On Error GoTo InitError
    ' Initialization code
    Exit Sub
InitError:
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_error_number_check() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    If Err.Number = 53 Then
        Resume Next
    ElseIf Err.Number = 5 Then
        Resume
    Else
        Resume ExitSub
    End If
ExitSub:
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_in_function() {
        let source = r"
Function Calculate() As Double
    On Error GoTo CalcError
    Calculate = x / y
    Exit Function
CalcError:
    Calculate = 0
    Resume Next
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_line_continuation() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume _
        Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_inline_if() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_transaction() {
        let source = r"
Sub ProcessTransaction()
    On Error GoTo TransError
    BeginTrans
    ' Transaction code
    CommitTrans
    Exit Sub
TransError:
    Rollback
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_goto_label() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    GoTo ProcessData
ProcessData:
    ' Code
    Exit Sub
ErrorHandler:
    Resume ProcessData
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_with_exit_statement() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    For i = 1 To 10
        If i = 5 Then Exit For
    Next i
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_file_operations() {
        let source = r#"
Sub ReadFile()
    On Error GoTo FileError
    Open "data.txt" For Input As #1
    Line Input #1, dataLine
    Close #1
    Exit Sub
FileError:
    If Err.Number = 53 Then
        Resume CreateFile
    Else
        Resume Next
    End If
CreateFile:
    ' Create file logic
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn resume_database_operations() {
        let source = r#"
Sub QueryDatabase()
    On Error GoTo DBError
    rs.Open "SELECT * FROM Users"
    Exit Sub
DBError:
    If Err.Number = 3021 Then
        Resume Next
    Else
        Resume CleanUp
    End If
CleanUp:
    rs.Close
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/exit_resume");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
