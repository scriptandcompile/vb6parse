//! Control flow statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 control flow statements:
//! - Conditional statements (If/Then/Else/ElseIf)
//! - Loop statements (Do/Loop, For/Next, For Each)
//! - Case statements (Select Case)
//! - Jump statements (GoTo, Exit, Label)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an If statement: If condition Then ... End If
    /// Handles both single-line and multi-line If statements
    ///
    /// IfStatement
    /// ├─ If keyword
    /// ├─ condition tokens
    /// ├─ Then keyword
    /// ├─ body tokens
    /// ├─ ElseIfClause (if present)
    /// │  ├─ ElseIf keyword
    /// │  ├─ condition tokens
    /// │  ├─ Then keyword
    /// │  └─ body tokens
    /// ├─ ElseClause (if present)
    /// │  ├─ Else keyword
    /// │  └─ body tokens
    /// ├─ End keyword
    /// └─ If keyword
    ///
    pub(super) fn parse_if_statement(&mut self) {
        self.builder.start_node(SyntaxKind::IfStatement.to_raw());

        // Consume "If" keyword
        self.consume_token();

        // Parse the conditional expression
        self.parse_conditional();

        // Consume "Then" if present
        if self.at_token(VB6Token::ThenKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Then
        self.consume_whitespace();

        // Check if this is a single-line If statement (has code on the same line after Then)
        let is_single_line = !self.at_token(VB6Token::Newline) && !self.is_at_end();

        if is_single_line {
            // Single-line If: parse the inline statement(s)
            // We parse until we hit a newline or reach a colon (which could indicate Else on same line)
            while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                // Check for inline Else (: Else or just Else on same line)
                if self.at_token(VB6Token::ElseKeyword) {
                    break;
                }

                // Try control flow statements first (Exit, GoTo, etc. can appear inline)
                if self.is_control_flow_keyword() {
                    self.parse_control_flow_statement();
                    continue;
                }

                // Try built-in statements
                if self.is_builtin_statement_keyword() {
                    self.parse_builtin_statement();
                    continue;
                }

                // Try to parse using centralized statement dispatcher
                if self.is_statement_keyword() {
                    self.parse_statement();
                    continue;
                }

                // Handle other inline constructs
                match self.current_token() {
                    Some(VB6Token::Whitespace)
                    | Some(VB6Token::EndOfLineComment)
                    | Some(VB6Token::RemComment) => {
                        self.consume_token();
                    }
                    Some(VB6Token::ColonOperator) => {
                        // Colon can separate statements or precede Else
                        self.consume_token();
                    }
                    _ => {
                        // Check if this looks like an assignment
                        if self.is_at_assignment() {
                            self.parse_assignment_statement();
                        } else {
                            // Consume as unknown
                            self.consume_token();
                        }
                    }
                }
            }

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        } else {
            // Multi-line If: consume newline after Then
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }

            // Parse body until "End If", "Else", or "ElseIf"
            self.parse_code_block(|parser| {
                (parser.at_token(VB6Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(VB6Token::IfKeyword))
                    || parser.at_token(VB6Token::ElseIfKeyword)
                    || parser.at_token(VB6Token::ElseKeyword)
            });

            // Handle ElseIf and Else clauses
            while !self.is_at_end() {
                if self.at_token(VB6Token::ElseIfKeyword) {
                    // Parse ElseIf clause
                    self.parse_elseif_clause();
                } else if self.at_token(VB6Token::ElseKeyword) {
                    // Parse Else clause
                    self.parse_else_clause();
                } else {
                    break;
                }
            }

            // Consume "End If" and trailing tokens
            if self.at_token(VB6Token::EndKeyword) {
                // Consume "End"
                self.consume_token();

                // Consume any whitespace between "End" and "If"
                self.consume_whitespace();

                // Consume "If"
                self.consume_token();

                // Consume until newline (including it)
                self.consume_until(VB6Token::Newline);

                // Consume the newline
                if self.at_token(VB6Token::Newline) {
                    self.consume_token();
                }
            }
        }

        self.builder.finish_node(); // IfStatement
    }

    /// Parse a Do...Loop statement.
    ///
    /// VB6 supports several forms of Do loops:
    /// - Do While condition...Loop
    /// - Do Until condition...Loop
    /// - Do...Loop While condition
    /// - Do...Loop Until condition
    /// - Do...Loop (infinite loop, requires Exit Do)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement)
    pub(super) fn parse_do_statement(&mut self) {
        // if we are now parsing a do statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DoStatement.to_raw());

        // Consume "Do" keyword
        self.consume_token();

        // Consume whitespace after Do
        self.consume_whitespace();

        // Check if we have While or Until after Do
        let has_top_condition =
            self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword);

        if has_top_condition {
            // Consume While or Until
            self.consume_token();

            // Parse condition - consume everything until newline
            self.parse_conditional();
        }

        // Consume newline after Do line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Loop"
        self.parse_code_block(|parser| parser.at_token(VB6Token::LoopKeyword));

        // Consume "Loop" keyword
        if self.at_token(VB6Token::LoopKeyword) {
            self.consume_token();

            // Consume whitespace after Loop
            self.consume_whitespace();

            // Check if we have While or Until after Loop
            if self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword) {
                // Consume While or Until
                self.consume_token();

                // Parse condition - consume everything until newline
                self.parse_conditional();
            }

            // Consume newline after Loop
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // DoStatement
    }

    /// Parse a For...Next statement.
    ///
    /// VB6 For...Next loop syntax:
    /// - For counter = start To end [Step step]...Next [counter]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement)
    pub(super) fn parse_for_statement(&mut self) {
        // if we are now parsing a for statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ForStatement.to_raw());

        // Consume "For" keyword
        self.consume_token();

        // Consume everything until "To" or newline
        // This includes: counter variable, "=", start value
        while !self.is_at_end()
            && !self.at_token(VB6Token::ToKeyword)
            && !self.at_token(VB6Token::Newline)
        {
            self.consume_token();
        }

        // Consume "To" keyword if present
        if self.at_token(VB6Token::ToKeyword) {
            self.consume_token();

            // Consume everything until "Step" or newline (the end value)
            while !self.is_at_end()
                && !self.at_token(VB6Token::StepKeyword)
                && !self.at_token(VB6Token::Newline)
            {
                self.consume_token();
            }

            // Consume "Step" keyword if present
            if self.at_token(VB6Token::StepKeyword) {
                self.consume_token();

                // Consume everything until newline (the step value)
                self.consume_until(VB6Token::Newline);
            }
        }

        // Consume newline after For line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Next"
        self.parse_code_block(|parser| parser.at_token(VB6Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(VB6Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional counter variable)
            self.consume_until(VB6Token::Newline);

            // Consume newline after Next
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // ForStatement
    }

    /// Parse a For Each...Next statement.
    ///
    /// VB6 For Each...Next loop syntax:
    /// - For Each element In collection...Next [element]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/for-eachnext-statement)
    pub(super) fn parse_for_each_statement(&mut self) {
        // if we are now parsing a for each statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ForEachStatement.to_raw());

        // Consume "For" keyword
        self.consume_token();

        // Consume whitespace
        self.consume_whitespace();

        // Consume "Each" keyword
        if self.at_token(VB6Token::EachKeyword) {
            self.consume_token();
        }

        // Consume everything until "In" or newline
        // This includes: element variable name and whitespace
        while !self.is_at_end()
            && !self.at_token(VB6Token::InKeyword)
            && !self.at_token(VB6Token::Newline)
        {
            self.consume_token();
        }

        // Consume "In" keyword if present
        if self.at_token(VB6Token::InKeyword) {
            self.consume_token();

            // Consume everything until newline (the collection)
            self.consume_until(VB6Token::Newline);
        }

        // Consume newline after For Each line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Next"
        self.parse_code_block(|parser| parser.at_token(VB6Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(VB6Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional element variable)
            self.consume_until(VB6Token::Newline);

            // Consume newline after Next
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // ForEachStatement
    }

    /// Parse a Select Case statement.
    ///
    /// Syntax:
    ///   Select Case testexpression
    ///     Case expression1
    ///       statements1
    ///     Case expression2
    ///       statements2
    ///     Case Else
    ///       statementsElse
    ///   End Select
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/select-case-statement)
    pub(super) fn parse_select_case_statement(&mut self) {
        // if we are now parsing a select case statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::SelectCaseStatement.to_raw());

        // Consume "Select" keyword
        self.consume_token();

        // Consume any whitespace between "Select" and "Case"
        self.consume_whitespace();

        // Consume "Case" keyword
        if self.at_token(VB6Token::CaseKeyword) {
            self.consume_token();
        }

        // Consume everything until newline (the test expression)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse Case clauses until "End Select"
        while !self.is_at_end() {
            // Check for "End Select"
            if self.at_token(VB6Token::EndKeyword)
                && self.peek_next_keyword() == Some(VB6Token::SelectKeyword)
            {
                break;
            }

            // Check for "Case" keyword
            if self.at_token(VB6Token::CaseKeyword) {
                // Check if this is "Case Else"
                let is_case_else = self.peek_next_keyword() == Some(VB6Token::ElseKeyword);

                if is_case_else {
                    // Parse Case Else clause
                    self.builder.start_node(SyntaxKind::CaseElseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume any whitespace between "Case" and "Else"
                    self.consume_whitespace();

                    // Consume "Else"
                    if self.at_token(VB6Token::ElseKeyword) {
                        self.consume_token();
                    }

                    // Consume until newline
                    self.consume_until(VB6Token::Newline);
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }

                    // Parse statements in Case Else until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseElseClause
                } else {
                    // Parse regular Case clause
                    self.builder.start_node(SyntaxKind::CaseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume the case expression(s) until newline
                    self.consume_until(VB6Token::Newline);
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }

                    // Parse statements in Case until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseClause
                }
            } else {
                // Consume whitespace, newlines, and comments
                self.consume_token();
            }
        }

        // Consume "End Select" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Select"
            self.consume_whitespace();

            // Consume "Select"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // SelectCaseStatement
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
    fn inline_if_then_goto() {
        let source = r#"
Sub Test()
    If x > 0 Then GoTo Positive
    Debug.Print "negative or zero"
Positive:
    Debug.Print "positive"
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("ThenKeyword"));
    }

    #[test]
    fn inline_if_then_call() {
        let source = r#"
Sub Test()
    If enabled Then Call DoSomething
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("CallStatement"));
    }

    #[test]
    fn inline_if_then_assignment() {
        let source = r#"
Sub Test()
    If x > 10 Then result = "large"
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("AssignmentStatement"));
    }

    #[test]
    fn inline_if_then_set() {
        let source = r#"
Sub Test()
    If obj Is Nothing Then Set obj = New MyClass
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("SetStatement"));
    }

    #[test]
    fn inline_if_then_exit() {
        let source = r#"
Sub Test()
    If errorOccurred Then Exit Sub
    Debug.Print "continuing"
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ExitKeyword"));
    }

    #[test]
    fn inline_if_then_multiple_statements() {
        let source = r#"
Sub Test()
    If condition Then x = 1: y = 2
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        let count = debug.matches("AssignmentStatement").count();
        assert_eq!(
            count, 2,
            "Expected 2 assignment statements separated by colon"
        );
    }

    #[test]
    fn inline_if_preserves_whitespace() {
        let source = r#"
Sub Test()
    If x > 0 Then GoTo Label1
Label1:
    x = 1
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("Whitespace"));
        assert!(debug.contains("Newline"));
    }

    #[test]
    fn inline_if_then_goto_with_comment() {
        let source = r#"
Sub Test()
    If x > 0 Then GoTo Positive ' go to positive case
Positive:
    result = x
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn inline_if_then_call_with_args() {
        let source = r#"
Sub Test()
    If ready Then Call Process(x, y, z)
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("CallStatement"));
    }

    #[test]
    fn inline_if_then_nested_calls() {
        let source = r#"
Sub Test()
    If value > 0 Then result = Calculate(value)
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("AssignmentStatement"));
    }

    #[test]
    fn inline_if_complex_condition() {
        let source = r#"
Sub Test()
    If x > 0 And y < 10 Then GoTo Valid
Valid:
    Process
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
    }

    #[test]
    fn inline_if_not_condition() {
        let source = r#"
Sub Test()
    If Not IsValid Then Exit Sub
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ExitKeyword"));
    }

    #[test]
    fn binary_conditional() {
        let code = r#"Sub Test()
    If x = 5 Then
    End If
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a CodeBlock
        let code_block = sub_statement
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::CodeBlock)
            .expect("SubStatement should contain a CodeBlock");

        // The CodeBlock should contain an IfStatement
        let if_statement = code_block
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::IfStatement)
            .expect("CodeBlock should contain an IfStatement");

        // The IfStatement should contain a BinaryConditional
        let binary_conditional = if_statement
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::BinaryConditional)
            .expect("IfStatement should contain a BinaryConditional");

        // Verify the BinaryConditional structure
        assert_eq!(binary_conditional.kind, SyntaxKind::BinaryConditional);
        assert!(
            !binary_conditional.is_token,
            "BinaryConditional should be a node, not a token"
        );

        // Verify the BinaryConditional contains the expected elements:
        // whitespace, identifier "x", whitespace, "=", whitespace, number "5", whitespace
        assert!(binary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
        assert!(binary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::EqualityOperator));
        assert!(binary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Number && c.text == "5"));
    }

    #[test]
    fn unary_conditional() {
        let code = r#"Sub Test()
    If Not isEmpty(x) Then
    End If
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a CodeBlock
        let code_block = sub_statement
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::CodeBlock)
            .expect("SubStatement should contain a CodeBlock");

        // The CodeBlock should contain an IfStatement
        let if_statement = code_block
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::IfStatement)
            .expect("CodeBlock should contain an IfStatement");

        // The IfStatement should contain a UnaryConditional
        let unary_conditional = if_statement
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::UnaryConditional)
            .expect("IfStatement should contain a UnaryConditional");

        // Verify the UnaryConditional structure
        assert_eq!(unary_conditional.kind, SyntaxKind::UnaryConditional);
        assert!(
            !unary_conditional.is_token,
            "UnaryConditional should be a node, not a token"
        );

        // Verify the UnaryConditional contains the expected elements:
        // whitespace, Not keyword, whitespace, identifier "isEmpty", parentheses, identifier "x", parentheses, whitespace
        assert!(unary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::NotKeyword));
        assert!(unary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "isEmpty"));
        assert!(unary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
        assert!(unary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::LeftParentheses));
        assert!(unary_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::RightParentheses));
    }

    #[test]
    fn nested_if_elseif_else() {
        let code = r#"Sub Test()
    If x > 0 Then
        If y > 0 Then
        ElseIf y < 0 Then
        Else
        End If
    ElseIf x < 0 Then
    Else
    End If
End Sub
"#;

        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a CodeBlock
        let code_block = sub_statement
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::CodeBlock)
            .expect("SubStatement should contain a CodeBlock");

        // Find the outer IfStatement in the CodeBlock
        let outer_if = code_block
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::IfStatement)
            .expect("CodeBlock should contain an outer IfStatement");

        // Verify outer If has a BinaryConditional (x > 0)
        let outer_conditional = outer_if
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::BinaryConditional)
            .expect("Outer IfStatement should contain a BinaryConditional");
        assert!(outer_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
        assert!(outer_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::GreaterThanOperator));

        // Find the CodeBlock inside the outer If
        let outer_code_block = outer_if
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::CodeBlock)
            .expect("Outer IfStatement should contain a CodeBlock");

        // Find the inner IfStatement (nested within the outer If's CodeBlock)
        let inner_if = outer_code_block
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::IfStatement)
            .expect("Outer CodeBlock should contain a nested IfStatement");

        // Verify inner If has a BinaryConditional (y > 0)
        let inner_conditional = inner_if
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::BinaryConditional)
            .expect("Inner IfStatement should contain a BinaryConditional");
        assert!(inner_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "y"));
        assert!(inner_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::GreaterThanOperator));

        // Verify inner If has ElseIf clause
        let inner_elseif = inner_if
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::ElseIfClause)
            .expect("Inner IfStatement should contain an ElseIfClause");

        // Verify inner ElseIf has a BinaryConditional (y < 0)
        let inner_elseif_conditional = inner_elseif
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::BinaryConditional)
            .expect("Inner ElseIfClause should contain a BinaryConditional");
        assert!(inner_elseif_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "y"));
        assert!(inner_elseif_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::LessThanOperator));

        // Verify inner If has Else clause
        assert!(
            inner_if
                .children
                .iter()
                .any(|child| child.kind == SyntaxKind::ElseClause),
            "Inner IfStatement should contain an ElseClause"
        );

        // Verify outer If has ElseIf clause
        let outer_elseif = outer_if
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::ElseIfClause)
            .expect("Outer IfStatement should contain an ElseIfClause");

        // Verify outer ElseIf has a BinaryConditional (x < 0)
        let outer_elseif_conditional = outer_elseif
            .children
            .iter()
            .find(|child| child.kind == SyntaxKind::BinaryConditional)
            .expect("Outer ElseIfClause should contain a BinaryConditional");
        assert!(outer_elseif_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
        assert!(outer_elseif_conditional
            .children
            .iter()
            .any(|c| c.kind == SyntaxKind::LessThanOperator));

        // Verify outer If has Else clause
        assert!(
            outer_if
                .children
                .iter()
                .any(|child| child.kind == SyntaxKind::ElseClause),
            "Outer IfStatement should contain an ElseClause"
        );
    }
}
