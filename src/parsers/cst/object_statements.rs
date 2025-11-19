//! Object statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 object manipulation statements:
//! - Call - Call a procedure
//! - Set - Assign object reference
//! - With - Execute statements on object
//!
//! Note: Array operations (ReDim) are in the array_statements module.
//! Note: Control flow statements (If, Do, For, Select Case, GoTo, Exit, Label)
//! are in the controlflow module.
//! Built-in system statements (AppActivate, Beep, ChDir, ChDrive) are in the
//! built_in_statements module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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
    pub(super) fn parse_call_statement(&mut self) {
        // if we are now parsing a call statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());

        // Consume "Call" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // CallStatement
    }

    /// Parse a Set statement.
    ///
    /// VB6 Set statement syntax:
    /// - Set objectVar = [New] objectExpression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
    pub(super) fn parse_set_statement(&mut self) {
        // if we are now parsing a set statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SetStatement.to_raw());

        // Consume "Set" keyword
        self.consume_token();

        // Consume everything until newline
        // This includes: variable, "=", [New], object expression
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // SetStatement
    }

    /// Parse a With statement.
    ///
    /// VB6 With statement syntax:
    /// - With object
    ///     .Property1 = value1
    ///     .Property2 = value2
    ///   End With
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)
    pub(super) fn parse_with_statement(&mut self) {
        // if we are now parsing a with statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WithStatement.to_raw());

        // Consume "With" keyword
        self.consume_token();

        // Consume everything until newline (the object expression)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the body until "End With"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::WithKeyword)
        });

        // Consume "End With" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "With"
            self.consume_whitespace();

            // Consume "With"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // WithStatement
    }

    /// Check if the current token is a statement keyword that parse_statement can handle.
    pub(super) fn is_statement_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::CallKeyword) | Some(VB6Token::SetKeyword) | Some(VB6Token::WithKeyword)
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types
    /// defined in this module (object manipulation statements).
    pub(super) fn parse_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::CallKeyword) => {
                self.parse_call_statement();
            }
            Some(VB6Token::SetKeyword) => {
                self.parse_set_statement();
            }
            Some(VB6Token::WithKeyword) => {
                self.parse_with_statement();
            }
            _ => {}
        }
    }

    /// Parse a Dim statement: Dim/Private/Public x As Type
    pub(super) fn parse_dim(&mut self) {
        // if we are now parsing a dim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DimStatement.to_raw());

        // Consume the keyword (Dim, Private, Public, etc.)
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // DimStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Call statement tests
    #[test]
    fn call_statement_simple() {
        let source = "Call MySubroutine()\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_with_arguments() {
        let source = "Call ProcessData(x, y, z)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("Call ProcessData"));
        assert!(cst.text().contains("x, y, z"));
    }

    #[test]
    fn call_statement_preserves_whitespace() {
        let source = "Call  MyFunction (  arg1 ,  arg2  )\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_in_sub() {
        let source = "Sub Main()\nCall DoSomething()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(sub_statement) = cst.child_at(0) {
            assert_eq!(sub_statement.kind, SyntaxKind::SubStatement);
            assert!(sub_statement.text.contains("Call DoSomething"));
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn call_statement_no_parentheses() {
        let source = "Call MySubroutine\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }

        assert_eq!(cst.text(), source);
    }

    #[test]
    fn multiple_call_statements() {
        let source = "Call First()\nCall Second()\nCall Third()\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 3);

        for i in 0..3 {
            if let Some(child) = cst.child_at(i) {
                assert_eq!(child.kind, SyntaxKind::CallStatement);
            }
        }

        assert!(cst.text().contains("Call First"));
        assert!(cst.text().contains("Call Second"));
        assert!(cst.text().contains("Call Third"));
    }

    #[test]
    fn call_statement_with_string_arguments() {
        let source = "Call ShowMessage(\"Hello, World!\")\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("\"Hello, World!\""));
    }

    #[test]
    fn call_statement_with_complex_expressions() {
        let source = "Call Calculate(x + y, z * 2, (a - b) / c)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);

        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }

        assert!(cst.text().contains("x + y"));
        assert!(cst.text().contains("z * 2"));
    }

    // Set statement tests
    #[test]
    fn set_statement_simple() {
        let source = r#"
Sub Test()
    Set obj = myObject
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("SetKeyword"));
    }

    #[test]
    fn set_statement_with_new() {
        let source = r#"
Sub Test()
    Set obj = New MyClass
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("NewKeyword"));
    }

    #[test]
    fn set_statement_to_nothing() {
        let source = r#"
Sub Test()
    Set obj = Nothing
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn set_statement_with_property_access() {
        let source = r#"
Sub Test()
    Set myObj.Property = otherObj
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn set_statement_with_collection_access() {
        let source = r#"
Sub Test()
    Set item = collection.Item(1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn multiple_set_statements() {
        let source = r#"
Sub Test()
    Set obj1 = New Class1
    Set obj2 = New Class2
    Set obj3 = Nothing
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let set_count = debug.matches("SetStatement").count();
        assert_eq!(set_count, 3);
    }

    #[test]
    fn set_statement_preserves_whitespace() {
        let source = r#"
Sub Test()
    Set   obj   =   New   MyClass
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn set_statement_in_function() {
        let source = r#"
Function GetObject() As Object
    Set GetObject = New MyClass
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetStatement"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn set_statement_at_module_level() {
        let source = r#"
Set globalObj = New MyClass
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("WithKeyword"));
        assert!(debug.contains("myObject"));
    }

    #[test]
    fn with_statement_nested_property() {
        let source = r#"
Sub Test()
    With obj.SubObject
        .Value = 42
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("Form1"));
    }

    #[test]
    fn with_statement_nested() {
        let source = r#"
Sub Test()
    With outer
        .Value1 = 1
        With .Inner
            .Value2 = 2
        End With
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("employee"));
    }

    #[test]
    fn with_statement_with_if() {
        let source = r#"
Sub Test()
    With obj
        If .IsValid Then
            .Process
        End If
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn with_statement_with_loop() {
        let source = r#"
Sub Test()
    With collection
        For i = 1 To .Count
            .Item(i).Process
        Next i
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("myArray"));
    }

    #[test]
    fn with_statement_function_result() {
        let source = r#"
Sub Test()
    With GetObject()
        .Property = value
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("GetObject"));
    }

    #[test]
    fn with_statement_empty() {
        let source = r#"
Sub Test()
    With obj
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn with_statement_sequential() {
        let source = r#"
Sub Test()
    With obj1
        .Value = 1
    End With
    With obj2
        .Value = 2
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WithStatement").count();
        assert_eq!(count, 2, "Expected 2 sequential With statements");
    }

    #[test]
    fn with_statement_preserves_whitespace() {
        let source = r#"
With obj
    .Property = value
End With
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn with_statement_new_object() {
        let source = r#"
Sub Test()
    With New MyClass
        .Initialize
        .Value = 42
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("NewKeyword"));
    }

    #[test]
    fn with_statement_at_module_level() {
        let source = r#"
With GlobalObject
    .Config = value
End With
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WithStatement"));
        assert!(debug.contains("GlobalObject"));
    }

    // Dim statement tests
    #[test]
    fn parse_dim_declaration() {
        let source = "Dim x As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn parse_private_declaration() {
        let source = "Private m_value As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Private m_value As Long\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn parse_public_declaration() {
        let source = "Public g_config As String\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Public g_config As String\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PublicKeyword"));
    }

    #[test]
    fn parse_multiple_variable_declaration() {
        let source = "Dim x, y, z As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x, y, z As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn parse_const_declaration() {
        let source = "Const MAX_SIZE = 100\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Const MAX_SIZE = 100\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("ConstKeyword"));
    }

    #[test]
    fn parse_private_const_declaration() {
        let source = "Private Const MODULE_NAME = \"MyModule\"\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Private Const MODULE_NAME = \"MyModule\"\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("ConstKeyword"));
    }

    #[test]
    fn parse_static_declaration() {
        let source = "Static counter As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Static counter As Long\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("StaticKeyword"));
    }
}
