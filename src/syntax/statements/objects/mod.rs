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
mod tests {
    use crate::assert_tree;
    use crate::*; // Call statement tests

    #[test]
    fn call_statement_simple() {
        let source = "Call MySubroutine()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("MySubroutine"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_with_arguments() {
        let source = "Call ProcessData(x, y, z)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                LeftParenthesis,
                Identifier ("x"),
                Comma,
                Whitespace,
                Identifier ("y"),
                Comma,
                Whitespace,
                Identifier ("z"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_preserves_whitespace() {
        let source = "Call  MyFunction (  arg1 ,  arg2  )\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("MyFunction"),
                Whitespace,
                LeftParenthesis,
                Whitespace,
                Identifier ("arg1"),
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("arg2"),
                Whitespace,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_in_sub() {
        let source = "Sub Main()\nCall DoSomething()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    CallStatement {
                        CallKeyword,
                        Whitespace,
                        Identifier ("DoSomething"),
                        LeftParenthesis,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_no_parentheses() {
        let source = "Call MySubroutine\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("MySubroutine"),
                Newline,
            },
        ]);
    }

    #[test]
    fn multiple_call_statements() {
        let source = "Call First()\nCall Second()\nCall Third()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("First"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("Second"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("Third"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_with_string_arguments() {
        let source = "Call ShowMessage(\"Hello, World!\")\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("ShowMessage"),
                LeftParenthesis,
                StringLiteral ("\"Hello, World!\""),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn call_statement_with_complex_expressions() {
        let source = "Call Calculate(x + y, z * 2, (a - b) / c)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("Calculate"),
                LeftParenthesis,
                Identifier ("x"),
                Whitespace,
                AdditionOperator,
                Whitespace,
                Identifier ("y"),
                Comma,
                Whitespace,
                Identifier ("z"),
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("2"),
                Comma,
                Whitespace,
                LeftParenthesis,
                Identifier ("a"),
                Whitespace,
                SubtractionOperator,
                Whitespace,
                Identifier ("b"),
                RightParenthesis,
                Whitespace,
                DivisionOperator,
                Whitespace,
                Identifier ("c"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    // Procedure call tests (without Call keyword)

    #[test]
    fn procedure_call_no_arguments() {
        let source = "InitializeRandomDNA\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                Identifier ("InitializeRandomDNA"),
                Newline,
            },
        ]);
    }

    #[test]
    fn procedure_call_with_parentheses() {
        let source = "DoSomething()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                Identifier ("DoSomething"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn procedure_call_with_arguments_no_parentheses() {
        let source = "MsgBox \"Hello\", vbInformation, \"Title\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Hello\""),
                Comma,
                Whitespace,
                Identifier ("vbInformation"),
                Comma,
                Whitespace,
                StringLiteral ("\"Title\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn procedure_call_with_arguments_with_parentheses() {
        let source = "ProcessData(x, y, z)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                Identifier ("ProcessData"),
                LeftParenthesis,
                Identifier ("x"),
                Comma,
                Whitespace,
                Identifier ("y"),
                Comma,
                Whitespace,
                Identifier ("z"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn multiple_procedure_calls_in_sub() {
        let source = "Sub Test()\nInitializeRandomDNA\nGetInitialSize\nGetInitialSpeed\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    CallStatement {
                        Identifier ("InitializeRandomDNA"),
                        Newline,
                    },
                    CallStatement {
                        Identifier ("GetInitialSize"),
                        Newline,
                    },
                    CallStatement {
                        Identifier ("GetInitialSpeed"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn procedure_call_preserves_whitespace() {
        let source = "MySub  arg1 ,  arg2\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            CallStatement {
                Identifier ("MySub"),
                Whitespace,
                Identifier ("arg1"),
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("arg2"),
                Newline,
            },
        ]);
    }

    #[test]
    fn procedure_call_vs_assignment() {
        // This should be an assignment, not a procedure call
        let source = "x = 5\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("x"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("5"),
                },
                Newline,
            },
        ]);
    }

    // Set statement tests

    #[test]
    fn set_statement_simple() {
        let source = r"
Sub Test()
    Set obj = myObject
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("myObject"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_with_new() {
        let source = r"
Sub Test()
    Set obj = New MyClass
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("MyClass"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_to_nothing() {
        let source = r"
Sub Test()
    Set obj = Nothing
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Nothing"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_with_property_access() {
        let source = r"
Sub Test()
    Set myObj.Property = otherObj
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("myObj"),
                        PeriodOperator,
                        PropertyKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("otherObj"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_with_function_call() {
        let source = r#"
Sub Test()
    Set result = GetObject("WinMgmts:")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("GetObject"),
                        LeftParenthesis,
                        StringLiteral ("\"WinMgmts:\""),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_with_collection_access() {
        let source = r"
Sub Test()
    Set item = collection.Item(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("item"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("collection"),
                        PeriodOperator,
                        Identifier ("Item"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj1"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Class1"),
                        Newline,
                    },
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj2"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Class2"),
                        Newline,
                    },
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj3"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Nothing"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    Set   obj   =   New   MyClass
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("MyClass"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_in_function() {
        let source = r"
Function GetObject() As Object
    Set GetObject = New MyClass
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetObject"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                ObjectKeyword,
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("GetObject"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("MyClass"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn set_statement_at_module_level() {
        let source = r"
Set globalObj = New MyClass
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SetStatement {
                SetKeyword,
                Whitespace,
                Identifier ("globalObj"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                NewKeyword,
                Whitespace,
                Identifier ("MyClass"),
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("myObject"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Property1"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"value\""),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Property2"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("123"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        PeriodOperator,
                        Identifier ("SubObject"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("42"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                        StatementList {
                            Whitespace,
                            Unknown,
                            CallStatement {
                                Identifier ("Show"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Caption"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Title\""),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("outer"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value1"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            WithStatement {
                                Whitespace,
                                WithKeyword,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("Inner"),
                                Newline,
                                StatementList {
                                    Whitespace,
                                    AssignmentStatement {
                                        IdentifierExpression {
                                            PeriodOperator,
                                        },
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("Value2"),
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2"),
                                            },
                                        },
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                EndKeyword,
                                Whitespace,
                                WithKeyword,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("employee"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("FirstName"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"John\""),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("LastName"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Doe\""),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Age"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("30"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Salary"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("50000"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                Identifier ("IsValid"),
                                Whitespace,
                                ThenKeyword,
                                Newline,
                            },
                            Whitespace,
                            Unknown,
                            CallStatement {
                                Identifier ("Process"),
                                Newline,
                            },
                            Whitespace,
                            Unknown,
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                IdentifierExpression {
                                    Newline,
                                },
                                Whitespace,
                                EndKeyword,
                                WithStatement {
                                    Whitespace,
                                    WithKeyword,
                                    Newline,
                                    StatementList {
                                        Unknown,
                                        Whitespace,
                                        Unknown,
                                        Newline,
                                    },
                                },
                            },
                        },
                    },
                },
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("collection"),
                        Newline,
                        StatementList {
                            ForStatement {
                                Whitespace,
                                ForKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                                Whitespace,
                                ToKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                Identifier ("Count"),
                                Newline,
                                StatementList {
                                    Whitespace,
                                    Unknown,
                                    CallStatement {
                                        Identifier ("Item"),
                                        LeftParenthesis,
                                        Identifier ("i"),
                                        RightParenthesis,
                                        PeriodOperator,
                                        Identifier ("Process"),
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                NextKeyword,
                                Whitespace,
                                Identifier ("i"),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("myArray"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        NameKeyword,
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Test\""),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("GetObject"),
                        LeftParenthesis,
                        RightParenthesis,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        PropertyKeyword,
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn with_statement_empty() {
        let source = r"
Sub Test()
    With obj
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Newline,
                        StatementList {
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj1"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj2"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn with_statement_preserves_whitespace() {
        let source = r"
With obj
    .Property = value
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("obj"),
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                PropertyKeyword,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("MyClass"),
                        Newline,
                        StatementList {
                            Whitespace,
                            Unknown,
                            CallStatement {
                                Identifier ("Initialize"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("42"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn with_statement_at_module_level() {
        let source = r"
With GlobalObject
    .Config = value
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("GlobalObject"),
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("Config"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    // RaiseEvent statement tests

    #[test]
    fn raiseevent_simple() {
        let source = r"
Sub Test()
    RaiseEvent DataReceived
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("DataReceived"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_single_argument() {
        let source = r"
Sub Test()
    RaiseEvent StatusChanged(newStatus)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("StatusChanged"),
                        LeftParenthesis,
                        Identifier ("newStatus"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_multiple_arguments() {
        let source = r"
Sub Test()
    RaiseEvent DataChanged(oldValue, newValue, timestamp)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("DataChanged"),
                        LeftParenthesis,
                        Identifier ("oldValue"),
                        Comma,
                        Whitespace,
                        Identifier ("newValue"),
                        Comma,
                        Whitespace,
                        Identifier ("timestamp"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_at_module_level() {
        let source = "RaiseEvent ProcessComplete\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            RaiseEventStatement {
                RaiseEventKeyword,
                Whitespace,
                Identifier ("ProcessComplete"),
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_preserves_whitespace() {
        let source = "    RaiseEvent    DataReady  (  value  )    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            RaiseEventStatement {
                RaiseEventKeyword,
                Whitespace,
                Identifier ("DataReady"),
                Whitespace,
                LeftParenthesis,
                Whitespace,
                Identifier ("value"),
                Whitespace,
                RightParenthesis,
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_comment() {
        let source = r"
Sub Test()
    RaiseEvent Updated ' Notify listeners
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("Updated"),
                        Whitespace,
                        EndOfLineComment,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("dataReady"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            RaiseEventStatement {
                                Whitespace,
                                RaiseEventKeyword,
                                Whitespace,
                                Identifier ("DataAvailable"),
                                LeftParenthesis,
                                Identifier ("data"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_inline_if() {
        let source = r"
Sub Test()
    If condition Then RaiseEvent EventFired
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("condition"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        RaiseEventStatement {
                            RaiseEventKeyword,
                            Whitespace,
                            Identifier ("EventFired"),
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                },
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                        StatementList {
                            RaiseEventStatement {
                                Whitespace,
                                RaiseEventKeyword,
                                Whitespace,
                                Identifier ("Progress"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_string_argument() {
        let source = r#"
Sub Test()
    RaiseEvent MessageSent("Hello, World!")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("MessageSent"),
                        LeftParenthesis,
                        StringLiteral ("\"Hello, World!\""),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_numeric_arguments() {
        let source = r"
Sub Test()
    RaiseEvent PositionChanged(100, 200)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("PositionChanged"),
                        LeftParenthesis,
                        IntegerLiteral ("100"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("200"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_object_property() {
        let source = r"
Sub Test()
    RaiseEvent ValueChanged(obj.Property)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("ValueChanged"),
                        LeftParenthesis,
                        Identifier ("obj"),
                        PeriodOperator,
                        PropertyKeyword,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_function_call() {
        let source = r"
Sub Test()
    RaiseEvent DataProcessed(ProcessData(input))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("DataProcessed"),
                        LeftParenthesis,
                        Identifier ("ProcessData"),
                        LeftParenthesis,
                        InputKeyword,
                        RightParenthesis,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("status"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("0"),
                            Newline,
                            StatementList {
                                RaiseEventStatement {
                                    Whitespace,
                                    RaiseEventKeyword,
                                    Whitespace,
                                    Identifier ("StatusIdle"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                RaiseEventStatement {
                                    Whitespace,
                                    RaiseEventKeyword,
                                    Whitespace,
                                    Identifier ("StatusBusy"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                RaiseEventStatement {
                                    Whitespace,
                                    RaiseEventKeyword,
                                    Whitespace,
                                    Identifier ("StatusError"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("BeforeUpdate"),
                        Newline,
                    },
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("Update"),
                        LeftParenthesis,
                        Identifier ("data"),
                        RightParenthesis,
                        Newline,
                    },
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("AfterUpdate"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_byref_argument() {
        let source = r"
Sub Test()
    RaiseEvent BeforeClose(Cancel)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("BeforeClose"),
                        LeftParenthesis,
                        Identifier ("Cancel"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("ErrorOccurred"),
                        LeftParenthesis,
                        Identifier ("Err"),
                        PeriodOperator,
                        Identifier ("Number"),
                        Comma,
                        Whitespace,
                        Identifier ("Err"),
                        PeriodOperator,
                        Identifier ("Description"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_array_element() {
        let source = r"
Sub Test()
    RaiseEvent ItemSelected(items(index))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("ItemSelected"),
                        LeftParenthesis,
                        Identifier ("items"),
                        LeftParenthesis,
                        Identifier ("index"),
                        RightParenthesis,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_expression() {
        let source = r"
Sub Test()
    RaiseEvent ProgressChanged((current / total) * 100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("ProgressChanged"),
                        LeftParenthesis,
                        LeftParenthesis,
                        Identifier ("current"),
                        Whitespace,
                        DivisionOperator,
                        Whitespace,
                        Identifier ("total"),
                        RightParenthesis,
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("100"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("processing"),
                        },
                        Newline,
                        StatementList {
                            RaiseEventStatement {
                                Whitespace,
                                RaiseEventKeyword,
                                Whitespace,
                                Identifier ("Processing"),
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("myObject"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Newline,
                            },
                            RaiseEventStatement {
                                Whitespace,
                                RaiseEventKeyword,
                                Whitespace,
                                Identifier ("ValueUpdated"),
                                LeftParenthesis,
                                PeriodOperator,
                                Identifier ("Value"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Whitespace,
                ClassKeyword,
                Newline,
            },
            PropertiesBlock {
                BeginKeyword,
                Newline,
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("MultiUse"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        SubtractionOperator,
                        IntegerLiteral ("1"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                EndKeyword,
                Newline,
            },
            EventStatement {
                PublicKeyword,
                Whitespace,
                EventKeyword,
                Whitespace,
                Identifier ("StatusChanged"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newStatus"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
            },
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("UpdateStatus"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("status"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("StatusChanged"),
                        LeftParenthesis,
                        Identifier ("status"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_boolean_argument() {
        let source = r"
Sub Test()
    RaiseEvent ValidationComplete(True)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("ValidationComplete"),
                        LeftParenthesis,
                        TrueKeyword,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_with_date_argument() {
        let source = r"
Sub Test()
    RaiseEvent DateChanged(#1/1/2025#)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("DateChanged"),
                        LeftParenthesis,
                        DateLiteral ("#1/1/2025#"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("shouldNotify"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            RaiseEventStatement {
                                Whitespace,
                                RaiseEventKeyword,
                                Whitespace,
                                Identifier ("Notification"),
                                LeftParenthesis,
                                Identifier ("message"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                RaiseEventStatement {
                                    Whitespace,
                                    RaiseEventKeyword,
                                    Whitespace,
                                    Identifier ("SilentUpdate"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn raiseevent_empty_parentheses() {
        let source = r"
Sub Test()
    RaiseEvent Complete()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RaiseEventStatement {
                        Whitespace,
                        RaiseEventKeyword,
                        Whitespace,
                        Identifier ("Complete"),
                        LeftParenthesis,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
