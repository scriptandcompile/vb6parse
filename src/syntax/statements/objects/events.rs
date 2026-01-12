//! `RaiseEvent` statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 `RaiseEvent` statements for firing custom events:
//! - `RaiseEvent` - Fire a custom event declared in a class or form
//!
//! # `RaiseEvent` Statement
//!
//! The `RaiseEvent` statement fires an event that has been declared within a class, form,
//! or document using the `Event` statement. Events can be raised with or without arguments.
//!
//! ## Syntax
//! ```vb
//! RaiseEvent eventName [(argumentList)]
//! ```
//!
//! ## Examples
//! ```vb
//! ' Event declaration (in class declarations)
//! Event DataReceived(data As String)
//! Event StatusChanged(oldStatus As Integer, newStatus As Integer)
//! Event ProcessComplete()
//!
//! ' Raising events
//! RaiseEvent ProcessComplete
//! RaiseEvent DataReceived("Test data")
//! RaiseEvent StatusChanged(0, 1)
//! ```
//!
//! ## Remarks
//! - Events must be declared with the `Event` statement before they can be raised
//! - `RaiseEvent` can only be used in the module where the event is declared
//! - Arguments passed must match the event declaration
//! - Events are consumed by objects that declare variables `WithEvents`
//! - Events cannot be raised recursively (no re-entrancy)
//! - Events raised in forms and controls are handled by the container
//!
//! ## Related Declarations
//! ```vb
//! ' Declaring events
//! Public Event StatusChange(status As Integer)
//!
//! ' Handling events (in consumer code)
//! Dim WithEvents obj As MyClass
//!
//! Private Sub obj_StatusChange(status As Integer)
//!     ' Handle the event
//! End Sub
//! ```
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raiseevent-statement)

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a `RaiseEvent` statement.
    ///
    /// VB6 `RaiseEvent` statement syntax:
    ///
    /// ```text
    /// RaiseEvent eventName [(argumentList)]
    /// ```
    ///
    /// ## Examples
    /// ```vb
    /// Sub ProcessData()
    ///     RaiseEvent DataReceived("Test data")
    ///     RaiseEvent StatusChanged(0, 1)
    ///     RaiseEvent ProcessComplete
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raiseevent-statement)
    pub(crate) fn parse_raiseevent_statement(&mut self) {
        // if we are now parsing a raiseevent statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::RaiseEventStatement.to_raw());
        self.consume_whitespace();

        // Consume "RaiseEvent" keyword
        self.consume_token();

        // Consume everything until newline (event name and arguments)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // RaiseEventStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn raiseevent_statement_simple() {
        let source = r"
Sub Test()
    RaiseEvent ProcessComplete
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_one_argument() {
        let source = r#"
Sub Test()
    RaiseEvent DataReceived("Test data")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_multiple_arguments() {
        let source = r"
Sub Test()
    RaiseEvent StatusChanged(0, 1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_parentheses_no_args() {
        let source = r"
Sub Test()
    RaiseEvent ProcessComplete()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_variable_argument() {
        let source = r"
Sub Test()
    Dim status As Integer
    status = 1
    RaiseEvent StatusUpdate(status)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_expression() {
        let source = r"
Sub Test()
    RaiseEvent ValueChanged(x + y)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_raiseevent_statements() {
        let source = r"
Sub Test()
    RaiseEvent Start
    RaiseEvent Progress(50)
    RaiseEvent Complete
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_if() {
        let source = r"
Sub Test()
    If condition Then
        RaiseEvent EventTriggered
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        RaiseEvent Progress(i)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_select_case() {
        let source = r"
Sub Test()
    Select Case status
        Case 0
            RaiseEvent Idle
        Case 1
            RaiseEvent Active
    End Select
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_with() {
        let source = r"
Sub Test()
    With myObject
        RaiseEvent .DataReady
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_error_handler() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    ' code
    Exit Sub
ErrorHandler:
    RaiseEvent ErrorOccurred(Err.Number)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    RaiseEvent   EventName   (   arg1   ,   arg2   )
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_named_arguments() {
        let source = r"
Sub Test()
    RaiseEvent DataUpdate(index:=1, value:=100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_in_property() {
        let source = r"
Property Let Value(v As Integer)
    mValue = v
    RaiseEvent ValueChanged(v)
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_object_property() {
        let source = r"
Sub Test()
    RaiseEvent PropertyUpdated(myObject.Property)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_function_call() {
        let source = r"
Sub Test()
    RaiseEvent ResultReady(Calculate(x, y))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_array_element() {
        let source = r"
Sub Test()
    RaiseEvent ItemSelected(items(index))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_byval_byref() {
        let source = r"
Sub Test()
    RaiseEvent DataProcessed(ByVal result, ByRef status)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_conditional_inline() {
        let source = r"
Sub Test()
    If ready Then RaiseEvent Ready Else RaiseEvent NotReady
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_constants() {
        let source = r"
Sub Test()
    RaiseEvent StatusChanged(STATUS_ACTIVE, True, 100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_nothing() {
        let source = r"
Sub Test()
    RaiseEvent ObjectReleased(Nothing)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_me() {
        let source = r"
Sub Test()
    RaiseEvent SourceChanged(Me)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_complex_expression() {
        let source = r"
Sub Test()
    RaiseEvent Calculation((x + y) * z / 2)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_at_module_level() {
        let source = r"
RaiseEvent GlobalEvent
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn raiseevent_statement_with_string_concatenation() {
        let source = r#"
Sub Test()
    RaiseEvent MessageSent("Hello " & userName & "!")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/events");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
