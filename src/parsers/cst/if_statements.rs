//! `If`/`Then`/`Else`/`ElseIf` statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 conditional statements:
//! - `If`/`Then`/`Else` statements (both single-line and multi-line)
//! - `ElseIf` clauses
//! - `Else` clauses

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse an `If` statement: `If` condition `Then` ... `End If`
    /// Handles both single-line and multi-line `If` statements
    ///
    /// `IfStatement`
    /// ├─ `If` keyword
    /// ├─ condition tokens
    /// ├─ `Then` keyword
    /// ├─ body tokens
    /// ├─ `ElseIfClause` (if present)
    /// │  ├─ `ElseIf` keyword
    /// │  ├─ condition tokens
    /// │  ├─ `Then` keyword
    /// │  └─ body tokens
    /// ├─ `ElseClause` (if present)
    /// │  ├─ `Else` keyword
    /// │  └─ body tokens
    /// ├─ `End` keyword
    /// └─ `If` keyword
    ///
    pub(crate) fn parse_if_statement(&mut self) {
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::IfStatement.to_raw());
        self.consume_whitespace();
        self.consume_token(); // If
        self.consume_whitespace();
        self.parse_expression();
        self.consume_whitespace();

        if self.at_token(Token::ThenKeyword) {
            self.consume_token();
        }
        self.consume_whitespace();

        // Check if single-line If
        let is_single_line = !self.at_token(Token::Newline) && !self.is_at_end();

        if is_single_line {
            // Single-line If: parse inline statements
            while !self.is_at_end() && !self.at_token(Token::Newline) {
                if self.at_token(Token::ElseKeyword) {
                    break;
                }

                if self.is_control_flow_keyword() {
                    self.parse_control_flow_statement();
                    continue;
                }
                if self.is_library_statement_keyword() {
                    self.parse_library_statement();
                    continue;
                }
                if self.is_variable_declaration_keyword() {
                    self.parse_array_statement();
                    continue;
                }
                if self.is_statement_keyword() {
                    self.parse_statement();
                    continue;
                }

                match self.current_token() {
                    Some(
                        Token::Whitespace
                        | Token::EndOfLineComment
                        | Token::RemComment
                        | Token::ColonOperator,
                    ) => {
                        self.consume_token();
                    }
                    _ => {
                        if self.at_token(Token::LetKeyword) {
                            self.parse_let_statement();
                        } else if self.is_at_assignment() {
                            self.parse_assignment_statement();
                        } else {
                            self.consume_token();
                        }
                    }
                }
            }

            if self.at_token(Token::Newline) {
                self.consume_token();
            }

            self.builder.finish_node(); // IfStatement
        } else {
            // Multi-line If: parse body and ElseIf/Else clauses
            if self.at_token(Token::Newline) {
                self.consume_token();
            }

            // Parse If body - the recursive call here is now safe because
            // parse_statement_list handles control flow iteratively
            self.parse_statement_list(|parser| {
                (parser.at_token(Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(Token::IfKeyword))
                    || parser.at_token(Token::ElseIfKeyword)
                    || parser.at_token(Token::ElseKeyword)
            });

            // Handle ElseIf and Else clauses
            while !self.is_at_end() {
                if self.at_token(Token::ElseIfKeyword) {
                    self.builder.start_node(SyntaxKind::ElseIfClause.to_raw());
                    self.consume_token(); // ElseIf
                    self.consume_whitespace();
                    self.parse_expression();
                    self.consume_whitespace();

                    if self.at_token(Token::ThenKeyword) {
                        self.consume_token();
                    }
                    self.consume_whitespace();

                    if self.at_token(Token::Newline) {
                        self.consume_token();
                    }

                    // Parse ElseIf body
                    self.parse_statement_list(|parser| {
                        parser.at_token(Token::ElseIfKeyword)
                            || parser.at_token(Token::ElseKeyword)
                            || (parser.at_token(Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(Token::IfKeyword))
                    });

                    self.builder.finish_node(); // ElseIfClause
                } else if self.at_token(Token::ElseKeyword) {
                    self.builder.start_node(SyntaxKind::ElseClause.to_raw());
                    self.consume_token(); // Else
                    self.consume_whitespace();

                    if self.at_token(Token::Newline) {
                        self.consume_token();
                    }

                    // Parse Else body
                    self.parse_statement_list(|parser| {
                        parser.at_token(Token::EndKeyword)
                            && parser.peek_next_keyword() == Some(Token::IfKeyword)
                    });

                    self.builder.finish_node(); // ElseClause
                } else {
                    break;
                }
            }

            // Consume "End If"
            if self.at_token(Token::EndKeyword) {
                self.consume_token();
                self.consume_whitespace();
                self.consume_token(); // If
                self.consume_until_after(Token::Newline);
            }

            self.builder.finish_node(); // IfStatement
        }
    }
}

#[cfg(test)]
mod tests {
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_call() {
        let source = r"
Sub Test()
    If enabled Then Call DoSomething
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_assignment() {
        let source = r#"
Sub Test()
    If x > 10 Then result = "large"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_set() {
        let source = r"
Sub Test()
    If obj Is Nothing Then Set obj = New MyClass
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_exit() {
        let source = r#"
Sub Test()
    If errorOccurred Then Exit Sub
    Debug.Print "continuing"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_multiple_statements() {
        let source = r"
Sub Test()
    If condition Then x = 1: y = 2
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_preserves_whitespace() {
        let source = r"
Sub Test()
    If x > 0 Then GoTo Label1
Label1:
    x = 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_goto_with_comment() {
        let source = r"
Sub Test()
    If x > 0 Then GoTo Positive ' go to positive case
Positive:
    result = x
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_call_with_args() {
        let source = r"
Sub Test()
    If ready Then Call Process(x, y, z)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_then_nested_calls() {
        let source = r"
Sub Test()
    If value > 0 Then result = Calculate(value)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_complex_condition() {
        let source = r"
Sub Test()
    If x > 0 And y < 10 Then GoTo Valid
Valid:
    Process
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn inline_if_not_condition() {
        let source = r"
Sub Test()
    If Not IsValid Then Exit Sub
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn binary_conditional() {
        let source = r"Sub Test()
    If x = 5 Then
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn unary_conditional() {
        let source = r"Sub Test()
    If Not isEmpty(x) Then
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[allow(clippy::too_many_lines)]
    #[test]
    fn nested_if_elseif_else() {
        let source = r"Sub Test()
    If x > 0 Then
        If y > 0 Then
        ElseIf y < 0 Then
        Else
        End If
    ElseIf x < 0 Then
    Else
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/if_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
