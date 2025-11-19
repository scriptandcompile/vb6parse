//! Control flow statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 control flow statements:
//! - Jump statements (GoTo, Exit, Label)
//!
//! Note: If/Then/Else/ElseIf statements are in the if_statements module.
//! Note: Select Case statements are in the select_statements module.
//! Note: For/Next and For Each/Next statements are in the for_statements module.
//! Note: Do/Loop statements are in the loop_statements module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // GotoStatement
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
}
