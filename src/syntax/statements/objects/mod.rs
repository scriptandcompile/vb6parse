//! Object statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 object manipulation statements:
//! - Call - Call a procedure
//! - `RaiseEvent` - Fire an event declared at module level
//! - Set - Assign object reference
//! - With - Execute statements on object
//!
//! Note: Variable declarations (Dim, `ReDim`)
//! are in the `variable_declarations` module.
//! Note: Control flow statements (If, Do, For, Select Case, `GoTo`, Exit, Label)
//! are in the controlflow module.
//! Built-in system statements (`AppActivate`, Beep, `ChDir`, `ChDrive`) are in the
//! library module.

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a Call statement:
    ///
    /// \[ Call \] name \[ argumentlist \]
    ///
    /// The Call statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Call        | Optional            | Indicates that a procedure is being called. The Call keyword is optional; if omitted, the procedure name is used directly. |
    /// | name        | Required            | Name of the procedure to be called; follows standard variable naming conventions. |
    /// | argumentlist| Optional            | List of arguments to be passed to the procedure. Arguments are enclosed in parentheses and separated by commas. |
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)
    pub(crate) fn parse_call_statement(&mut self) {
        // if we are now parsing a call statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());
        self.consume_whitespace();

        // Consume "Call" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // CallStatement
    }

    /// Parse a procedure call without the Call keyword.
    /// In VB6, you can call a Sub procedure without using the Call keyword:
    /// - `MySub arg1, arg2` instead of `Call MySub(arg1, arg2)`
    /// - `MySub` (no arguments)
    pub(crate) fn parse_procedure_call(&mut self) {
        // if we are now parsing a procedure call, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());
        self.consume_whitespace();

        // Consume everything until newline (procedure name and arguments)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // CallStatement
    }

    /// Check if the current position is at a procedure call (without Call keyword).
    /// This is true if we have an identifier that's not followed by an assignment operator.
    /// In VB6, procedure calls can appear as:
    /// - `MySub` (no arguments)
    /// - `MySub arg1, arg2` (arguments without parentheses)
    /// - `MySub(arg1, arg2)` (arguments with parentheses)
    pub(crate) fn is_at_procedure_call(&self) -> bool {
        // Must start with an identifier or keyword used as identifier
        // BUT exclude keywords that have structural meaning and can't be procedure names
        if self.at_token(Token::Identifier) {
            // Identifiers are OK
        } else if self.at_keyword() {
            // Some keywords should never be treated as procedure calls
            // These are structural keywords that have special parsing rules
            if let Some(
                Token::EndKeyword
                | Token::ExitKeyword
                | Token::LoopKeyword
                | Token::NextKeyword
                | Token::WendKeyword
                | Token::ElseKeyword
                | Token::ElseIfKeyword
                | Token::CaseKeyword
                | Token::IfKeyword
                | Token::ThenKeyword
                | Token::SelectKeyword
                | Token::DoKeyword
                | Token::WhileKeyword
                | Token::UntilKeyword
                | Token::ForKeyword
                | Token::ToKeyword
                | Token::StepKeyword
                | Token::SubKeyword
                | Token::FunctionKeyword
                | Token::PropertyKeyword
                | Token::WithKeyword
                | Token::ReturnKeyword
                | Token::ResumeKeyword,
            ) = self.current_token()
            {
                return false;
            }
        } else {
            return false;
        }

        // Look ahead to see if there's an assignment operator
        // If there's an =, it's an assignment, not a procedure call
        for (_text, token) in self.tokens.iter().skip(self.pos) {
            match token {
                Token::Newline | Token::EndOfLineComment | Token::RemComment => {
                    // Reached end of line without finding assignment - this is a procedure call
                    return true;
                }
                Token::EqualityOperator => {
                    // Found = operator - this is an assignment, not a procedure call
                    return false;
                }
                // Procedure calls can have various tokens before newline
                Token::Identifier
                | Token::LeftParenthesis
                | Token::RightParenthesis
                | Token::Comma
                | Token::PeriodOperator
                | Token::StringLiteral
                | Token::IntegerLiteral
                | Token::LongLiteral
                | Token::SingleLiteral
                | Token::DoubleLiteral => {
                    // These can all appear in procedure calls, continue looking
                }
                // If it's a keyword, it could be an argument
                _ if token.is_keyword() => {
                    // Keywords can be used as arguments (e.g., True, False, Nothing)
                }
                // Whitespace or Anything else could indicate it's not a simple procedure call
                _ => {}
            }
        }

        false
    }

    /// Parse a `RaiseEvent` statement.
    ///
    /// VB6 `RaiseEvent` statement syntax:
    /// - `RaiseEvent` eventname [(argumentlist)]
    ///
    /// Fires an event declared at module level within a class, form, or document.
    ///
    /// The `RaiseEvent` statement syntax has these parts:
    ///
    /// | Part         | Description |
    /// |--------------|-------------|
    /// | eventname    | Required. Name of the event to fire. |
    /// | argumentlist | Optional. Comma-delimited list of variables, arrays, or expressions. The argumentlist must match the parameters defined in the Event declaration. |
    ///
    /// Remarks:
    /// - If the event has no arguments, don't include the parentheses.
    /// - `RaiseEvent` can only be used to fire events declared in the same class or form module.
    /// - Events can't be raised within a standard module.
    /// - When an event is raised, all procedures connected to that event are executed.
    /// - Events can have `ByVal` and `ByRef` arguments like normal procedures.
    /// - Events with arguments can be cancelled by the event handler if declared with a Cancel parameter.
    /// - `RaiseEvent` can only fire events that are explicitly declared with the Event statement in the same module.
    ///
    /// Examples:
    /// ```vb
    /// ' In a class module
    /// Public Event DataReceived(ByVal data As String)
    /// Public Event StatusChanged(ByVal oldStatus As Integer, ByVal newStatus As Integer)
    /// Public Event ProcessComplete()
    ///
    /// Private Sub ProcessData()
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

    /// Parse a Set statement.
    ///
    /// VB6 Set statement syntax:
    /// - Set objectVar = [New] objectExpression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
    pub(crate) fn parse_set_statement(&mut self) {
        // if we are now parsing a set statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SetStatement.to_raw());
        self.consume_whitespace();

        // Consume "Set" keyword
        self.consume_token();

        // Consume everything until newline
        // This includes: variable, "=", [New], object expression
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // SetStatement
    }

    /// Parse a With statement.
    ///
    /// VB6 With statement syntax:
    ///
    /// ```vb
    /// With object
    ///     .Property1 = value1
    ///     .Property2 = value2
    /// End With
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)
    pub(crate) fn parse_with_statement(&mut self) {
        // if we are now parsing a with statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WithStatement.to_raw());
        self.consume_whitespace();

        // Consume "With" keyword
        self.consume_token();

        // Consume everything until newline (the object expression)
        self.consume_until_after(Token::Newline);

        // Parse the body until "End With"
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::WithKeyword)
        });

        // Consume "End With" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "With"
            self.consume_whitespace();

            // Consume "With"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // WithStatement
    }

    /// Check if the current token is a statement keyword that `parse_statement` can handle.
    /// Checks both current position and next non-whitespace token.
    pub(crate) fn is_statement_keyword(&self) -> bool {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        matches!(
            token,
            Some(
                Token::CallKeyword
                    | Token::RaiseEventKeyword
                    | Token::SetKeyword
                    | Token::WithKeyword
            )
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types
    /// defined in this module (object manipulation statements).
    pub(crate) fn parse_statement(&mut self) {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        match token {
            Some(Token::CallKeyword) => {
                self.parse_call_statement();
            }
            Some(Token::RaiseEventKeyword) => {
                self.parse_raiseevent_statement();
            }
            Some(Token::SetKeyword) => {
                self.parse_set_statement();
            }
            Some(Token::WithKeyword) => {
                self.parse_with_statement();
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Call statement tests
    #[test]
    fn call_statement_simple() {
        let source = "Call MySubroutine()\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_with_arguments() {
        let source = "Call ProcessData(x, y, z)\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("Call ProcessData"));
        assert!(cst.text().contains("x, y, z"));
    }

    #[test]
    fn call_statement_preserves_whitespace() {
        let source = "Call  MyFunction (  arg1 ,  arg2  )\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_in_sub() {
        let source = "Sub Main()\nCall DoSomething()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(sub_statement) = cst.child_at(0) {
            assert_eq!(sub_statement.kind(), SyntaxKind::SubStatement);
            assert!(sub_statement.text().contains("Call DoSomething"));
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_no_parentheses() {
        let source = "Call MySubroutine\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn multiple_call_statements() {
        let source = "Call First()\nCall Second()\nCall Third()\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 3);

        for i in 0..3 {
            if let Some(child) = cst.child_at(i) {
                assert_eq!(child.kind(), SyntaxKind::CallStatement);
            }
        }

        assert!(cst.text().contains("Call First"));
        assert!(cst.text().contains("Call Second"));
        assert!(cst.text().contains("Call Third"));
    }

    #[test]
    fn call_statement_with_string_arguments() {
        let source = "Call ShowMessage(\"Hello, World!\")\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("\"Hello, World!\""));
    }

    #[test]
    fn call_statement_with_complex_expressions() {
        let source = "Call Calculate(x + y, z * 2, (a - b) / c)\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("x + y"));
        assert!(cst.text().contains("z * 2"));
    }

    // Procedure call tests (without Call keyword)
    #[test]
    fn procedure_call_no_arguments() {
        let source = "InitializeRandomDNA\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn procedure_call_with_parentheses() {
        let source = "DoSomething()\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn procedure_call_with_arguments_no_parentheses() {
        let source = "MsgBox \"Hello\", vbInformation, \"Title\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("MsgBox"));
    }

    #[test]
    fn procedure_call_with_arguments_with_parentheses() {
        let source = "ProcessData(x, y, z)\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("ProcessData"));
    }

    #[test]
    fn multiple_procedure_calls_in_sub() {
        let source = "Sub Test()\nInitializeRandomDNA\nGetInitialSize\nGetInitialSpeed\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(sub_statement) = cst.child_at(0) {
            assert_eq!(sub_statement.kind(), SyntaxKind::SubStatement);

            // Check that the debug tree contains CallStatements
            let debug = cst.debug_tree();
            assert!(debug.contains("CallStatement"));

            // Count CallStatements - there should be at least 3
            let call_count = debug
                .lines()
                .filter(|line| line.contains("CallStatement@"))
                .count();
            assert!(
                call_count >= 3,
                "Expected at least 3 CallStatements, found {call_count}"
            );
        }
    }

    #[test]
    fn procedure_call_preserves_whitespace() {
        let source = "MySub  arg1 ,  arg2\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn procedure_call_vs_assignment() {
        // This should be an assignment, not a procedure call
        let source = "x = 5\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind(), SyntaxKind::AssignmentStatement);
        }
    }

    // Set statement tests
    #[test]
    fn set_statement_simple() {
        let source = r"
Sub Test()
    Set obj = myObject
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("SetKeyword"));
    }

    #[test]
    fn set_statement_with_new() {
        let source = r"
Sub Test()
    Set obj = New MyClass
End Sub
";

        let mut source_stream = SourceStream::new("test.bas", source);
        let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("NewKeyword"));
    }

    #[test]
    fn set_statement_to_nothing() {
        let source = r"
Sub Test()
    Set obj = Nothing
End Sub
";

        let mut source_stream = SourceStream::new("test.bas", source);
        let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn set_statement_with_property_access() {
        let source = r"
Sub Test()
    Set myObj.Property = otherObj
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("PeriodOperator"));
    }

    #[test]
    fn set_statement_with_function_call() {
        let source = r#"
Sub Test()
    Set result = GetObject("WinMgmts:")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn set_statement_with_collection_access() {
        let source = r"
Sub Test()
    Set item = collection.Item(1)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn multiple_set_statements() {
        let source = r"
Sub Test()
    Set obj1 = New Class1
    Set obj2 = New Class2
    Set obj3 = Nothing
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let set_count = debug.matches("SetStatement").count();
        assert_eq!(set_count, 3);
    }

    #[test]
    fn set_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    Set   obj   =   New   MyClass
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn set_statement_in_function() {
        let source = r"
Function GetObject() As Object
    Set GetObject = New MyClass
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn set_statement_at_module_level() {
        let source = r"
Set globalObj = New MyClass
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    // With statement tests
    #[test]
    fn with_statement_simple() {
        let source = r#"
Sub Test()
    With myObject
        .Property1 = "value"
        .Property2 = 123
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("WithKeyword"));
        assert!(debug.contains("myObject"));
    }

    #[test]
    fn with_statement_nested_property() {
        let source = r"
Sub Test()
    With obj.SubObject
        .Value = 42
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("SubObject"));
    }

    #[test]
    fn with_statement_method_call() {
        let source = r#"
Sub Test()
    With Form1
        .Show
        .Caption = "Title"
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("Form1"));
    }

    #[test]
    fn with_statement_nested() {
        let source = r"
Sub Test()
    With outer
        .Value1 = 1
        With .Inner
            .Value2 = 2
        End With
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WithStatement").count();
        assert_eq!(count, 2, "Expected 2 With statements (nested)");
    }

    #[test]
    fn with_statement_multiple_properties() {
        let source = r#"
Sub Test()
    With employee
        .FirstName = "John"
        .LastName = "Doe"
        .Age = 30
        .Salary = 50000
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("employee"));
    }

    #[test]
    fn with_statement_with_if() {
        let source = r"
Sub Test()
    With obj
        If .IsValid Then
            .Process
        End If
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn with_statement_with_loop() {
        let source = r"
Sub Test()
    With collection
        For i = 1 To .Count
            .Item(i).Process
        Next i
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn with_statement_array_access() {
        let source = r#"
Sub Test()
    With myArray(5)
        .Name = "Test"
        .Value = 100
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("myArray"));
    }

    #[test]
    fn with_statement_function_result() {
        let source = r"
Sub Test()
    With GetObject()
        .Property = value
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("GetObject"));
    }

    #[test]
    fn with_statement_empty() {
        let source = r"
Sub Test()
    With obj
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn with_statement_sequential() {
        let source = r"
Sub Test()
    With obj1
        .Value = 1
    End With
    With obj2
        .Value = 2
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WithStatement").count();
        assert_eq!(count, 2, "Expected 2 sequential With statements");
    }

    #[test]
    fn with_statement_preserves_whitespace() {
        let source = r"
With obj
    .Property = value
End With
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn with_statement_new_object() {
        let source = r"
Sub Test()
    With New MyClass
        .Initialize
        .Value = 42
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("NewKeyword"));
    }

    #[test]
    fn with_statement_at_module_level() {
        let source = r"
With GlobalObject
    .Config = value
End With
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("GlobalObject"));
    }

    // RaiseEvent statement tests
    #[test]
    fn raiseevent_simple() {
        let source = r"
Sub Test()
    RaiseEvent DataReceived
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("RaiseEventKeyword"));
        assert!(debug.contains("DataReceived"));
    }

    #[test]
    fn raiseevent_with_single_argument() {
        let source = r"
Sub Test()
    RaiseEvent StatusChanged(newStatus)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("StatusChanged"));
        assert!(debug.contains("newStatus"));
    }

    #[test]
    fn raiseevent_with_multiple_arguments() {
        let source = r"
Sub Test()
    RaiseEvent DataChanged(oldValue, newValue, timestamp)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("oldValue"));
        assert!(debug.contains("newValue"));
        assert!(debug.contains("timestamp"));
    }

    #[test]
    fn raiseevent_at_module_level() {
        let source = "RaiseEvent ProcessComplete\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
    }

    #[test]
    fn raiseevent_preserves_whitespace() {
        let source = "    RaiseEvent    DataReady  (  value  )    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    RaiseEvent    DataReady  (  value  )    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
    }

    #[test]
    fn raiseevent_with_comment() {
        let source = r"
Sub Test()
    RaiseEvent Updated ' Notify listeners
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn raiseevent_in_if_statement() {
        let source = r"
Sub Test()
    If dataReady Then
        RaiseEvent DataAvailable(data)
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn raiseevent_inline_if() {
        let source = r"
Sub Test()
    If condition Then RaiseEvent EventFired
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
    }

    #[test]
    fn raiseevent_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        RaiseEvent Progress(i)
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn raiseevent_with_string_argument() {
        let source = r#"
Sub Test()
    RaiseEvent MessageSent("Hello, World!")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("Hello"));
    }

    #[test]
    fn raiseevent_with_numeric_arguments() {
        let source = r"
Sub Test()
    RaiseEvent PositionChanged(100, 200)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("100"));
        assert!(debug.contains("200"));
    }

    #[test]
    fn raiseevent_with_object_property() {
        let source = r"
Sub Test()
    RaiseEvent ValueChanged(obj.Property)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("obj"));
        assert!(debug.contains("Property"));
    }

    #[test]
    fn raiseevent_with_function_call() {
        let source = r"
Sub Test()
    RaiseEvent DataProcessed(ProcessData(input))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("ProcessData"));
    }

    #[test]
    fn raiseevent_in_select_case() {
        let source = r"
Sub Test()
    Select Case status
        Case 0
            RaiseEvent StatusIdle
        Case 1
            RaiseEvent StatusBusy
        Case 2
            RaiseEvent StatusError
    End Select
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let raiseevent_count = debug.matches("RaiseEventStatement").count();
        assert_eq!(raiseevent_count, 3);
    }

    #[test]
    fn raiseevent_multiple_in_sequence() {
        let source = r"
Sub Test()
    RaiseEvent BeforeUpdate
    RaiseEvent Update(data)
    RaiseEvent AfterUpdate
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let raiseevent_count = debug.matches("RaiseEventStatement").count();
        assert_eq!(raiseevent_count, 3);
    }

    #[test]
    fn raiseevent_with_byref_argument() {
        let source = r"
Sub Test()
    RaiseEvent BeforeClose(Cancel)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("Cancel"));
    }

    #[test]
    fn raiseevent_in_error_handler() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    RaiseEvent ErrorOccurred(Err.Number, Err.Description)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("ErrorOccurred"));
    }

    #[test]
    fn raiseevent_with_array_element() {
        let source = r"
Sub Test()
    RaiseEvent ItemSelected(items(index))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("items"));
    }

    #[test]
    fn raiseevent_with_expression() {
        let source = r"
Sub Test()
    RaiseEvent ProgressChanged((current / total) * 100)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("current"));
        assert!(debug.contains("total"));
    }

    #[test]
    fn raiseevent_in_do_loop() {
        let source = r"
Sub Test()
    Do While processing
        RaiseEvent Processing
    Loop
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn raiseevent_in_with_block() {
        let source = r"
Sub Test()
    With myObject
        .Value = 100
        RaiseEvent ValueUpdated(.Value)
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn raiseevent_in_class_module() {
        let source = r"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Public Event StatusChanged(ByVal newStatus As String)

Public Sub UpdateStatus(ByVal status As String)
    RaiseEvent StatusChanged(status)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("EventStatement"));
    }

    #[test]
    fn raiseevent_with_boolean_argument() {
        let source = r"
Sub Test()
    RaiseEvent ValidationComplete(True)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("True"));
    }

    #[test]
    fn raiseevent_with_date_argument() {
        let source = r"
Sub Test()
    RaiseEvent DateChanged(#1/1/2025#)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
    }

    #[test]
    fn raiseevent_conditional_firing() {
        let source = r"
Sub Test()
    If shouldNotify Then
        RaiseEvent Notification(message)
    Else
        RaiseEvent SilentUpdate
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let raiseevent_count = debug.matches("RaiseEventStatement").count();
        assert_eq!(raiseevent_count, 2);
    }

    #[test]
    fn raiseevent_empty_parentheses() {
        let source = r"
Sub Test()
    RaiseEvent Complete()
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RaiseEventStatement"));
        assert!(debug.contains("Complete"));
    }
}
