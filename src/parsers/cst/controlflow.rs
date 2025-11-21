//! Control flow statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 control flow statements:
//! - Jump statements (GoTo, GoSub, Return, Resume, Exit, Label)
//!
//! Note: If/Then/Else/ElseIf statements are in the if_statements module.
//! Note: Select Case statements are in the select_statements module.
//! Note: For/Next and For Each/Next statements are in the for_statements module.
//! Note: Do/Loop statements are in the loop_statements module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a GoSub statement.
    ///
    /// VB6 GoSub statement syntax:
    /// - GoSub label
    ///
    /// Branches to and returns from a subroutine within a procedure.
    ///
    /// The GoSub...Return statement syntax has these parts:
    ///
    /// | Part   | Description |
    /// |--------|-------------|
    /// | label  | Required. A line label or line number. |
    ///
    /// Remarks:
    /// - You can use GoSub and Return anywhere in a procedure, but GoSub and the corresponding Return statement must be in the same procedure.
    /// - A subroutine can contain more than one Return statement, but the first one encountered causes the flow of execution to branch back to the statement immediately following the most recently executed GoSub statement.
    /// - You can't enter or exit Sub procedures with GoSub...Return.
    /// - Using GoSub and Return is considered obsolete. Modern VB6 code should use Sub or Function procedures instead.
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
    pub(super) fn parse_gosub_statement(&mut self) {
        // if we are now parsing a gosub statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::GoSubStatement.to_raw());

        // Consume "GoSub" keyword
        self.consume_token();

        // Consume everything until newline (the label name)
        self.consume_until_after(VB6Token::Newline);

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
    /// - Return must be used with GoSub to return to the statement following the GoSub call.
    /// - You can use GoSub and Return anywhere in a procedure, but GoSub and the corresponding Return statement must be in the same procedure.
    /// - A subroutine can contain more than one Return statement, but the first one encountered causes the flow of execution to branch back to the statement immediately following the most recently executed GoSub statement.
    /// - Using GoSub and Return is considered obsolete. Modern VB6 code should use Sub or Function procedures instead.
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
    pub(super) fn parse_return_statement(&mut self) {
        // if we are now parsing a return statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ReturnStatement.to_raw());

        // Consume "Return" keyword
        self.consume_token();

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ReturnStatement
    }

    /// Parse a GoTo statement.
    ///
    /// Syntax:
    ///   GoTo label
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(super) fn parse_goto_statement(&mut self) {
        // if we are now parsing a goto statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::GotoStatement.to_raw());

        // Consume "GoTo" keyword
        self.consume_token();

        // Consume everything until newline (the label name)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // GotoStatement
    }

    /// Parse a Resume statement.
    ///
    /// VB6 Resume statement syntax:
    /// - Resume
    /// - Resume Next
    /// - Resume label
    ///
    /// Resumes execution after an error-handling routine is finished.
    ///
    /// # Syntax
    ///
    /// The Resume statement has these forms:
    ///
    /// | Form | Description |
    /// |------|-------------|
    /// | Resume | If the error occurred in the same procedure as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the statement that last called out of the procedure containing the error-handling routine. |
    /// | Resume Next | If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or On Error Resume Next statement). |
    /// | Resume label | Execution resumes at the line specified by the label argument. The label argument can be a line label or line number. |
    ///
    /// # Remarks
    ///
    /// - The Resume statement can be used only in an error-handling routine.
    /// - Using Resume without specifying a label causes execution to resume at the statement that caused the error.
    /// - Resume Next is useful when you want to continue execution despite an error.
    /// - Resume label is useful when you want to continue execution at a specific location after handling an error.
    /// - If you use a Resume statement anywhere except in an error-handling routine, an error occurs.
    /// - Resume cannot be used in any procedure that contains an On Error Resume Next statement.
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
    pub(super) fn parse_resume_statement(&mut self) {
        // if we are now parsing a resume statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ResumeStatement.to_raw());

        // Consume "Resume" keyword
        self.consume_token();

        // Consume everything until newline (Next keyword or label)
        self.consume_until_after(VB6Token::Newline);

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
    pub(super) fn parse_exit_statement(&mut self) {
        // if we are now parsing an exit statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ExitStatement.to_raw());

        // Consume "Exit" keyword
        self.consume_token();

        // Consume whitespace after Exit
        self.consume_whitespace();

        // Consume the exit type (Do, For, Function, Property, Sub)
        if self.at_token(VB6Token::DoKeyword)
            || self.at_token(VB6Token::ForKeyword)
            || self.at_token(VB6Token::FunctionKeyword)
            || self.at_token(VB6Token::PropertyKeyword)
            || self.at_token(VB6Token::SubKeyword)
        {
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ExitStatement
    }

    /// Parse a label statement.
    ///
    /// VB6 label syntax:
    /// - LabelName:
    ///
    /// Labels are used as targets for GoTo and GoSub statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(super) fn parse_label_statement(&mut self) {
        // if we are now parsing a label statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::LabelStatement.to_raw());

        // Consume the label identifier
        self.consume_token();

        // Consume optional whitespace
        self.consume_whitespace();

        // Consume the colon
        if self.at_token(VB6Token::ColonOperator) {
            self.consume_token();
        }

        // Consume the newline if present
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LabelStatement
    }

    /// Check if the current position is at a label.
    /// A label is an identifier or number followed by a colon.
    pub(super) fn is_at_label(&self) -> bool {
        let next_token_is_colon = matches!(self.peek_next_token(), Some(VB6Token::ColonOperator));

        if next_token_is_colon == false {
            return false;
        }

        // If we are not parsing the header, then some keywords are valid identifiers (like "Begin")
        // TODO: Consider adding a list of keywords that can be used as labels.
        // TODO: Also consider modifying tokenizer to recognize when inside header to more easily identify Identifiers vs header only keywords.
        if !self.parsing_header && matches!(self.current_token(), Some(VB6Token::BeginKeyword)) {
            return true;
        }

        self.is_identifier() || self.is_number()
    }

    /// Parse an On Error statement.
    ///
    /// VB6 On Error statement syntax:
    /// - On Error GoTo label
    /// - On Error GoTo 0
    /// - On Error Resume Next
    ///
    /// Enables an error-handling routine and specifies the location of the routine within a procedure.
    ///
    /// The On Error statement syntax has these forms:
    ///
    /// | Form | Description |
    /// |------|-------------|
    /// | On Error GoTo line | Enables the error-handling routine that starts at line. The line argument is any line label or line number. If a run-time error occurs, control branches to line, making the error handler active. |
    /// | On Error Resume Next | Specifies that when a run-time error occurs, control goes to the statement immediately following the statement where the error occurred, and execution continues from that point. |
    /// | On Error GoTo 0 | Disables any enabled error handler in the current procedure. |
    ///
    /// Remarks:
    /// - If you don't use an On Error statement, any run-time error that occurs is fatal; that is, an error message is displayed and execution stops.
    /// - An "enabled" error handler is one that is turned on by an On Error statement. An "active" error handler is an enabled handler that is in the process of handling an error.
    /// - If an error occurs while an error handler is active (between the occurrence of the error and a Resume, Exit Sub, Exit Function, or Exit Property statement), the current procedure's error handler can't handle the error.
    /// - Control returns to the calling procedure. If the calling procedure has an enabled error handler, it is activated to handle the error.
    /// - If the calling procedure's error handler is also active, control passes back through previous calling procedures until an enabled, but inactive, error handler is found.
    /// - If no inactive, enabled error handler is found, the error is fatal at the point at which it actually occurred.
    /// - Each time the error handler passes control back to a calling procedure, that procedure becomes the current procedure. Once an error is handled in any procedure, execution resumes in the current procedure at the point designated by the Resume statement.
    ///
    /// Examples:
    /// ```vb
    /// Sub Test()
    ///     On Error GoTo ErrorHandler
    ///     ' Code that might cause an error
    ///     Exit Sub
    /// ErrorHandler:
    ///     MsgBox "An error occurred: " & Err.Description
    /// End Sub
    ///
    /// Sub Test2()
    ///     On Error Resume Next
    ///     ' Code continues even if errors occur
    ///     MkDir "C:\Temp"  ' Won't stop if directory exists
    /// End Sub
    ///
    /// Sub Test3()
    ///     On Error GoTo 0  ' Disable error handling
    ///     ' Normal error behavior
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)
    pub(super) fn parse_on_error_statement(&mut self) {
        // if we are now parsing an on error statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::OnErrorStatement.to_raw());

        // Consume "On" keyword
        self.consume_token();

        // Consume "Error" keyword
        if self.at_token(VB6Token::ErrorKeyword) {
            self.consume_token();
        }

        // Consume everything until newline (GoTo label, Resume Next, GoTo 0, etc.)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // OnErrorStatement
    }

    /// Parse an On GoTo statement.
    ///
    /// VB6 On GoTo statement syntax:
    /// - On expression GoTo label1[, label2, ...]
    ///
    /// Branches to one of several specified labels, depending on the value of an expression.
    ///
    /// The On...GoTo statement syntax has these parts:
    ///
    /// | Part | Description |
    /// |------|-------------|
    /// | expression | Required. Any numeric expression that evaluates to a whole number between 0 and 255, inclusive. If expression is any number other than a whole number, it is rounded before it is evaluated. |
    /// | labellist | Required. List of line labels or line numbers separated by commas. |
    ///
    /// Remarks:
    /// - The value of expression determines which line is branched to in the list of labels. If the value of expression is less than 1 or greater than the number of items in the list, one of the following results occurs:
    ///   - If expression equals 0, execution continues with the statement following On...GoTo.
    ///   - If expression is greater than the number of labels in the list, execution continues with the statement following On...GoTo.
    ///   - If expression is negative or greater than 255, an error occurs.
    /// - The On...GoTo statement is useful for branching to one of several different labels based on a value.
    /// - Using On...GoTo is considered obsolete. Modern VB6 code should use Select Case instead.
    ///
    /// Examples:
    /// ```vb
    /// Sub Test()
    ///     Dim choice As Integer
    ///     choice = 2
    ///     On choice GoTo Label1, Label2, Label3
    ///     Exit Sub
    /// Label1:
    ///     MsgBox "Choice 1"
    ///     Exit Sub
    /// Label2:
    ///     MsgBox "Choice 2"
    ///     Exit Sub
    /// Label3:
    ///     MsgBox "Choice 3"
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ongoto-and-ongosub-statements)
    pub(super) fn parse_on_goto_statement(&mut self) {
        // if we are now parsing an on goto statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::OnGoToStatement.to_raw());

        // Consume "On" keyword
        self.consume_token();

        // Consume everything until newline (expression GoTo labels)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // OnGoToStatement
    }

    /// Parse an On GoSub statement.
    ///
    /// VB6 On GoSub statement syntax:
    /// - On expression GoSub label1[, label2, ...]
    ///
    /// Branches to one of several specified subroutines, depending on the value of an expression.
    ///
    /// The On...GoSub statement syntax has these parts:
    ///
    /// | Part | Description |
    /// |------|-------------|
    /// | expression | Required. Any numeric expression that evaluates to a whole number between 0 and 255, inclusive. If expression is any number other than a whole number, it is rounded before it is evaluated. |
    /// | labellist | Required. List of line labels or line numbers separated by commas. |
    ///
    /// Remarks:
    /// - The value of expression determines which subroutine is called in the list of labels. If the value of expression is less than 1 or greater than the number of items in the list, one of the following results occurs:
    ///   - If expression equals 0, execution continues with the statement following On...GoSub.
    ///   - If expression is greater than the number of labels in the list, execution continues with the statement following On...GoSub.
    ///   - If expression is negative or greater than 255, an error occurs.
    /// - The On...GoSub statement is useful for branching to one of several different subroutines based on a value.
    /// - Each subroutine must end with a Return statement to return to the statement following the On...GoSub.
    /// - Using On...GoSub is considered obsolete. Modern VB6 code should use Select Case with Sub procedure calls instead.
    ///
    /// Examples:
    /// ```vb
    /// Sub Test()
    ///     Dim menuChoice As Integer
    ///     menuChoice = 1
    ///     On menuChoice GoSub Menu1, Menu2, Menu3
    ///     Exit Sub
    /// Menu1:
    ///     MsgBox "Menu 1 selected"
    ///     Return
    /// Menu2:
    ///     MsgBox "Menu 2 selected"
    ///     Return
    /// Menu3:
    ///     MsgBox "Menu 3 selected"
    ///     Return
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ongoto-and-ongosub-statements)
    pub(super) fn parse_on_gosub_statement(&mut self) {
        // if we are now parsing an on gosub statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::OnGoSubStatement.to_raw());

        // Consume "On" keyword
        self.consume_token();

        // Consume everything until newline (expression GoSub labels)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // OnGoSubStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn exit_do() {
        let source = r#"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("DoKeyword"));
    }

    #[test]
    fn exit_for() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 5 Then Exit For
    Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("ForKeyword"));
    }

    #[test]
    fn exit_function() {
        let source = r#"
Function Test() As Integer
    If x = 0 Then
        Exit Function
    End If
    Test = 42
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("FunctionKeyword"));
    }

    #[test]
    fn exit_sub() {
        let source = r#"
Sub Test()
    If x = 0 Then Exit Sub
    Debug.Print "x is not zero"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("SubKeyword"));
    }

    #[test]
    fn exit_property() {
        let source = r#"
Property Set Callback(ByRef newObj As InterPress)
    Set mCallback = newObj
    Exit Property
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("PropertyKeyword"));
    }

    #[test]
    fn multiple_exit_statements() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 3 Then Exit For
        If i = 7 Then Exit Sub
    Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Should have two ExitStatements
        let exit_count = debug.matches("ExitStatement").count();
        assert_eq!(exit_count, 2);
    }

    #[test]
    fn exit_in_nested_loops() {
        let source = r#"
Sub Test()
    Do While x < 100
        For i = 1 To 10
            If i = 5 Then Exit For
        Next
        If x > 50 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let exit_count = debug.matches("ExitStatement").count();
        assert_eq!(exit_count, 2);
    }

    #[test]
    fn exit_preserves_whitespace() {
        let source = r#"
Sub Test()
    Exit   Sub
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ExitStatement"));
        // Check that whitespace is preserved
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn inline_exit_in_if_statement() {
        let source = r#"
Function Test(x As Integer) As Integer
    If x = 0 Then Exit Function
    Test = x * 2
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
        assert!(debug.contains("FunctionKeyword"));
    }

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("GotoKeyword"));
        assert!(debug.contains("ErrorHandler"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("GotoKeyword"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("LargeValue"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("GotoStatement").count();
        assert_eq!(count, 3, "Expected 3 GoTo statements");
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Note: "On Error GoTo" is a special case that may be parsed differently
        // This test just ensures we can handle the basic GoTo part
        assert!(debug.contains("GotoKeyword"));
    }

    #[test]
    fn goto_statement_forward_jump() {
        let source = r#"
Sub Test()
    x = 1
    GoTo SkipCode
    x = 2
    x = 3
SkipCode:
    x = 4
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("SkipCode"));
    }

    #[test]
    fn goto_statement_backward_jump() {
        let source = r#"
Sub Test()
StartLoop:
    counter = counter + 1
    If counter < 10 Then
        GoTo StartLoop
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("StartLoop"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("SelectCaseStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("ForStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("Error_Handler"));
    }

    #[test]
    fn goto_statement_preserves_whitespace() {
        let source = r#"
Sub Test()
    GoTo MyLabel
MyLabel:
    x = 1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("Whitespace"));
        assert!(debug.contains("Newline"));
    }

    #[test]
    fn goto_statement_same_line_as_then() {
        let source = r#"
Sub Test()
    If condition Then
        GoTo Handler
    End If
Handler:
    result = True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("Handler"));
    }

    #[test]
    fn goto_statement_exit_cleanup() {
        let source = r#"
Sub Test()
    On Error GoTo Cleanup
    ' do work
    Exit Sub
Cleanup:
    ' cleanup code
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GotoKeyword"));
    }

    #[test]
    fn label_simple() {
        let source = r#"
Sub Test()
    MyLabel:
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("MyLabel"));
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

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("ErrorHandler"));
    }

    #[test]
    fn label_with_underscore() {
        let source = r#"
Sub Test()
Error_Handler:
    MsgBox "Error"
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("Error_Handler"));
    }

    #[test]
    fn label_at_module_level() {
        let source = r#"
Sub Test()
StartHere:
    x = 1
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("StartHere"));
    }

    #[test]
    fn label_multiple() {
        let source = r#"
Sub Test()
Start:
    x = 1
Middle:
    y = 2
End_Label:
    z = 3
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("LabelStatement").count();
        assert_eq!(count, 3, "Expected 3 label statements");
    }

    #[test]
    fn label_with_space_after_colon() {
        let source = r#"
Sub Test()
MyLabel: x = 1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("MyLabel"));
    }

    #[test]
    fn label_error_handler_pattern() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Some code
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("ErrorHandler"));
    }

    #[test]
    fn label_with_numbers() {
        let source = r#"
Sub Test()
Label123:
    x = 1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("Label123"));
    }

    #[test]
    fn label_cleanup_pattern() {
        let source = r#"
Sub Test()
    GoTo Cleanup
Cleanup:
    Set obj = Nothing
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("Cleanup"));
    }

    #[test]
    fn label_preserves_whitespace() {
        let source = "MyLabel:";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("MyLabel"));
        assert!(debug.contains("ColonOperator"));
    }

    #[test]
    fn label_in_function() {
        let source = r#"
Function Calculate() As Integer
Start:
    Calculate = 42
End Function
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("Start"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn label_mixed_case() {
        let source = r#"
Sub Test()
MyErrorHandler:
    MsgBox "Error"
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LabelStatement"));
        assert!(debug.contains("MyErrorHandler"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("GoSubKeyword"));
        assert!(debug.contains("ErrorHandler"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("GoSubKeyword"));
    }

    #[test]
    fn gosub_in_if_statement() {
        let source = r#"
Sub Test()
    If x > 0 Then
        GoSub ProcessPositive
    End If
    Exit Sub
ProcessPositive:
    y = y + 1
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("ProcessPositive"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let gosub_count = debug.matches("GoSubStatement").count();
        assert_eq!(gosub_count, 2);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let gosub_count = debug.matches("GoSubStatement").count();
        assert_eq!(gosub_count, 2);
    }

    #[test]
    fn gosub_preserves_whitespace() {
        let source = r#"
Sub Test()
    GoSub   MyLabel
MyLabel:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn gosub_with_comment() {
        let source = r#"
Sub Test()
    GoSub Cleanup ' Call cleanup routine
Cleanup:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn gosub_inline_if() {
        let source = r#"
Sub Test()
    If needsInit Then GoSub Initialize
    Exit Sub
Initialize:
    x = 0
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("Initialize"));
    }

    #[test]
    fn gosub_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        GoSub ProcessItem
    Next i
    Exit Sub
ProcessItem:
    Debug.Print i
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn gosub_in_select_case() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("SelectCaseStatement"));
    }

    #[test]
    fn gosub_with_underscore_label() {
        let source = r#"
Sub Test()
    GoSub Error_Handler
    Exit Sub
Error_Handler:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("Error_Handler"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("DoWork"));
    }

    // Return statement tests
    #[test]
    fn return_simple() {
        let source = r#"
Sub Test()
    GoSub SubRoutine
    Exit Sub
SubRoutine:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("ReturnKeyword"));
    }

    #[test]
    fn return_multiple() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let return_count = debug.matches("ReturnStatement").count();
        assert_eq!(return_count, 2);
    }

    #[test]
    fn return_in_if_statement() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn return_inline_if() {
        let source = r#"
Sub Test()
    GoSub Validate
    Exit Sub
Validate:
    If invalid Then Return
    DoSomething
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
    }

    #[test]
    fn return_with_comment() {
        let source = r#"
Sub Test()
    GoSub Cleanup
    Exit Sub
Cleanup:
    Set obj = Nothing
    Return ' Exit subroutine
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn return_preserves_whitespace() {
        let source = r#"
Sub Test()
    GoSub Sub1
    Exit Sub
Sub1:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("Whitespace"));
        assert!(debug.contains("Newline"));
    }

    #[test]
    fn return_in_select_case() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("SelectCaseStatement"));
    }

    #[test]
    fn return_in_loop() {
        let source = r#"
Sub Test()
    GoSub FindValue
    Exit Sub
FindValue:
    For i = 1 To 10
        If arr(i) = target Then Return
    Next i
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn gosub_return_complete_example() {
        let source = r#"
Sub Main()
    x = 10
    GoSub DoubleValue
    Debug.Print x
    Exit Sub
DoubleValue:
    x = x * 2
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("ReturnStatement"));
        assert!(debug.contains("LabelStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let gosub_count = debug.matches("GoSubStatement").count();
        let return_count = debug.matches("ReturnStatement").count();
        assert_eq!(gosub_count, 2);
        assert_eq!(return_count, 2);
    }

    #[test]
    fn return_at_module_level() {
        let source = r#"
Public Sub TestReturn()
    GoSub Handler
    Exit Sub
Handler:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReturnStatement"));
    }

    #[test]
    fn gosub_return_error_pattern() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GoSubStatement"));
        assert!(debug.contains("ReturnStatement"));
    }

    // On Error statement tests
    #[test]
    fn on_error_goto_label() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Code that might error
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("OnKeyword"));
        assert!(debug.contains("ErrorKeyword"));
    }

    #[test]
    fn on_error_resume_next() {
        let source = r#"
Sub Test()
    On Error Resume Next
    MkDir "C:\Temp"
    MkDir "C:\Data"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("OnKeyword"));
        assert!(debug.contains("ErrorKeyword"));
        assert!(debug.contains("ResumeKeyword"));
        assert!(debug.contains("NextKeyword"));
    }

    #[test]
    fn on_error_goto_0() {
        let source = r#"
Sub Test()
    On Error GoTo 0
    ' Error handling disabled
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("OnKeyword"));
        assert!(debug.contains("ErrorKeyword"));
    }

    #[test]
    fn on_error_at_module_level() {
        let source = r#"On Error GoTo ErrorHandler"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn on_error_with_whitespace() {
        let source = "On    Error    GoTo    Handler\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn on_error_with_comment() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler ' Setup error handling
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn on_error_in_if_statement() {
        let source = r#"
Sub Test()
    If needsErrorHandling Then
        On Error GoTo ErrorHandler
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn on_error_inline_if() {
        let source = r#"
Sub Test()
    If debug Then On Error GoTo 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn multiple_on_error_statements() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Do something
    On Error GoTo 0
    ' Disable error handling
    On Error Resume Next
    ' Continue on error
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("OnErrorStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn on_error_complete_pattern() {
        let source = r#"
Sub ProcessFile(filePath As String)
    On Error GoTo ErrorHandler
    
    Open filePath For Input As #1
    ' Process file
    Close #1
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    End If
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("LabelStatement"));
    }

    #[test]
    fn on_error_numeric_label() {
        let source = r#"
Sub Test()
    On Error GoTo 100
    ' Code
    Exit Sub
100:
    MsgBox "Error"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn on_error_nested_procedures() {
        let source = r#"
Sub Outer()
    On Error GoTo OuterError
    Inner
    Exit Sub
OuterError:
    MsgBox "Outer error"
End Sub

Sub Inner()
    On Error Resume Next
    ' Code
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("OnErrorStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn on_goto_simple() {
        let source = r#"
Sub Test()
    Dim choice As Integer
    choice = 2
    On choice GoTo Label1, Label2, Label3
    Exit Sub
Label1:
    MsgBox "Choice 1"
    Exit Sub
Label2:
    MsgBox "Choice 2"
    Exit Sub
Label3:
    MsgBox "Choice 3"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("OnKeyword"));
    }

    #[test]
    fn on_goto_with_expression() {
        let source = r#"
Sub Test()
    On x + 1 GoTo First, Second, Third
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("GotoKeyword"));
    }

    #[test]
    fn on_goto_single_label() {
        let source = r#"
Sub Test()
    On errorCode GoTo ErrorHandler
ErrorHandler:
    MsgBox "Error"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("ErrorHandler"));
    }

    #[test]
    fn on_goto_numeric_labels() {
        let source = r#"
Sub Test()
    On choice GoTo 100, 200, 300
    Exit Sub
100:
    x = 1
    Exit Sub
200:
    x = 2
    Exit Sub
300:
    x = 3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
    }

    #[test]
    fn on_goto_preserves_whitespace() {
        let source = r#"
Sub Test()
    On   choice   GoTo   Label1  ,  Label2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn on_goto_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        On menuChoice GoTo Menu1, Menu2
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn on_goto_inline_if() {
        let source = r#"
Sub Test()
    If useMenu Then On choice GoTo Menu1, Menu2, Menu3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
    }

    #[test]
    fn on_goto_at_module_level() {
        let source = r#"On choice GoTo Label1, Label2"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
    }

    #[test]
    fn on_goto_with_comment() {
        let source = r#"
Sub Test()
    On choice GoTo Label1, Label2 ' Branch based on choice
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn on_goto_complex_expression() {
        let source = r#"
Sub Test()
    On (x * 2) + 1 GoTo First, Second, Third, Fourth
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
    }

    #[test]
    fn multiple_on_goto_statements() {
        let source = r#"
Sub Test()
    On choice1 GoTo A1, A2, A3
    On choice2 GoTo B1, B2, B3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("OnGoToStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn on_goto_many_labels() {
        let source = r#"
Sub Test()
    On choice GoTo L1, L2, L3, L4, L5, L6, L7, L8, L9, L10
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoToStatement"));
    }

    #[test]
    fn on_gosub_simple() {
        let source = r#"
Sub Test()
    Dim menuChoice As Integer
    menuChoice = 1
    On menuChoice GoSub Menu1, Menu2, Menu3
    Exit Sub
Menu1:
    MsgBox "Menu 1 selected"
    Return
Menu2:
    MsgBox "Menu 2 selected"
    Return
Menu3:
    MsgBox "Menu 3 selected"
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("OnKeyword"));
    }

    #[test]
    fn on_gosub_with_expression() {
        let source = r#"
Sub Test()
    On x Mod 3 GoSub First, Second, Third
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("GoSubKeyword"));
    }

    #[test]
    fn on_gosub_single_label() {
        let source = r#"
Sub Test()
    On flag GoSub Handler
    Exit Sub
Handler:
    Debug.Print "Called"
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("Handler"));
    }

    #[test]
    fn on_gosub_numeric_labels() {
        let source = r#"
Sub Test()
    On choice GoSub 100, 200, 300
    Exit Sub
100:
    Debug.Print "100"
    Return
200:
    Debug.Print "200"
    Return
300:
    Debug.Print "300"
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
    }

    #[test]
    fn on_gosub_preserves_whitespace() {
        let source = r#"
Sub Test()
    On   choice   GoSub   Sub1  ,  Sub2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn on_gosub_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        On choice GoSub Handler1, Handler2
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn on_gosub_inline_if() {
        let source = r#"
Sub Test()
    If needsProcessing Then On choice GoSub Process1, Process2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
    }

    #[test]
    fn on_gosub_at_module_level() {
        let source = r#"On choice GoSub Sub1, Sub2"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
    }

    #[test]
    fn on_gosub_with_comment() {
        let source = r#"
Sub Test()
    On choice GoSub Sub1, Sub2 ' Call appropriate subroutine
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn on_gosub_nested() {
        let source = r#"
Sub Test()
    On choice1 GoSub Level1A, Level1B
    Exit Sub
Level1A:
    On choice2 GoSub Level2A, Level2B
    Return
Level1B:
    Return
Level2A:
    Return
Level2B:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("OnGoSubStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn on_gosub_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        On i Mod 3 GoSub Handler1, Handler2, Handler3
    Next i
    Exit Sub
Handler1:
    Return
Handler2:
    Return
Handler3:
    Return
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnGoSubStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn multiple_on_gosub_statements() {
        let source = r#"
Sub Test()
    On choice1 GoSub A1, A2, A3
    On choice2 GoSub B1, B2, B3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("OnGoSubStatement").count();
        assert_eq!(count, 2);
    }

    // Resume statement tests
    #[test]
    fn resume_simple() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("ResumeKeyword"));
    }

    #[test]
    fn resume_next() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    x = 1 / 0
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("ResumeKeyword"));
        assert!(debug.contains("NextKeyword"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("CleanUp"));
    }

    #[test]
    fn resume_with_line_number() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume 100
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("100"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 2);
    }

    #[test]
    fn resume_with_comment() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume Next ' Continue after error
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn resume_preserves_whitespace() {
        let source = "    Resume    Next    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Resume    Next    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_in_nested_error_handler() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 2);
    }

    #[test]
    fn resume_at_module_level() {
        let source = "Resume Next\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_with_select_case() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 3);
    }

    #[test]
    fn resume_in_loop() {
        let source = r#"
Sub ProcessFiles()
    On Error GoTo ErrorHandler
    For i = 1 To 10
        ProcessFile i
    Next i
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("ForStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 4);
    }

    #[test]
    fn resume_with_do_loop() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Do
        ' Code
    Loop
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn multiple_resume_statements() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 3);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OnErrorStatement"));
        assert!(debug.contains("ResumeKeyword"));
    }

    #[test]
    fn resume_in_class_module() {
        let source = r#"
Private Sub Class_Initialize()
    On Error GoTo InitError
    ' Initialization code
    Exit Sub
InitError:
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_with_error_number_check() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 3);
    }

    #[test]
    fn resume_in_function() {
        let source = r#"
Function Calculate() As Double
    On Error GoTo CalcError
    Calculate = x / y
    Exit Function
CalcError:
    Calculate = 0
    Resume Next
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn resume_with_line_continuation() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Resume _
        Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_inline_if() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_with_transaction() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
    }

    #[test]
    fn resume_with_goto_label() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    GoTo ProcessData
ProcessData:
    ' Code
    Exit Sub
ErrorHandler:
    Resume ProcessData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("GotoStatement"));
    }

    #[test]
    fn resume_with_exit_statement() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    For i = 1 To 10
        If i = 5 Then Exit For
    Next i
    Exit Sub
ErrorHandler:
    Resume Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResumeStatement"));
        assert!(debug.contains("ExitStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 2);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let resume_count = debug.matches("ResumeStatement").count();
        assert_eq!(resume_count, 2);
    }
}
