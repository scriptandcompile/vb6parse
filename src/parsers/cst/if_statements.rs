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
    pub(super) fn parse_if_statement(&mut self) {
        self.builder.start_node(SyntaxKind::IfStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "If" keyword
        self.consume_token();

        // Skip any leading whitespace
        self.consume_whitespace();

        // Parse the conditional expression
        self.parse_expression();

        // Consume whitespace before Then
        self.consume_whitespace();

        // Consume "Then" if present
        if self.at_token(Token::ThenKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Then
        self.consume_whitespace();

        // Check if this is a single-line If statement (has code on the same line after Then)
        let is_single_line = !self.at_token(Token::Newline) && !self.is_at_end();

        if is_single_line {
            // Single-line If: parse the inline statement(s)
            // We parse until we hit a newline or reach a colon (which could indicate Else on same line)
            while !self.is_at_end() && !self.at_token(Token::Newline) {
                // Check for inline Else (: Else or just Else on same line)
                if self.at_token(Token::ElseKeyword) {
                    break;
                }

                // Try control flow statements first (Exit, GoTo, etc. can appear inline)
                if self.is_control_flow_keyword() {
                    self.parse_control_flow_statement();
                    continue;
                }

                // Try built-in library statements
                if self.is_library_statement_keyword() {
                    self.parse_library_statement();
                    continue;
                }

                // Try variable declaration statements
                if self.is_variable_declaration_keyword() {
                    self.parse_array_statement();
                    continue;
                }

                // Try to parse using centralized statement dispatcher
                if self.is_statement_keyword() {
                    self.parse_statement();
                    continue;
                }

                // Handle other inline constructs
                match self.current_token() {
                    Some(Token::Whitespace | Token::EndOfLineComment | Token::RemComment) => {
                        self.consume_token();
                    }
                    Some(Token::ColonOperator) => {
                        // Colon can separate statements or precede Else
                        self.consume_token();
                    }
                    _ => {
                        // Check for Let statement (optional assignment keyword)
                        if self.at_token(Token::LetKeyword) {
                            self.parse_let_statement();
                        // Check if this looks like an assignment
                        } else if self.is_at_assignment() {
                            self.parse_assignment_statement();
                        } else {
                            // Consume as unknown
                            self.consume_token();
                        }
                    }
                }
            }

            // Consume the newline
            if self.at_token(Token::Newline) {
                self.consume_token();
            }
        } else {
            // Multi-line If: consume newline after Then
            if self.at_token(Token::Newline) {
                self.consume_token();
            }

            // Parse body until "End If", "Else", or "ElseIf"
            self.parse_statement_list(|parser| {
                (parser.at_token(Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(Token::IfKeyword))
                    || parser.at_token(Token::ElseIfKeyword)
                    || parser.at_token(Token::ElseKeyword)
            });

            // Handle ElseIf and Else clauses
            while !self.is_at_end() {
                if self.at_token(Token::ElseIfKeyword) {
                    // Parse ElseIf clause
                    self.parse_elseif_clause();
                } else if self.at_token(Token::ElseKeyword) {
                    // Parse Else clause
                    self.parse_else_clause();
                } else {
                    break;
                }
            }

            // Consume "End If" and trailing tokens
            if self.at_token(Token::EndKeyword) {
                // Consume "End"
                self.consume_token();

                // Consume any whitespace between "End" and "If"
                self.consume_whitespace();

                // Consume "If"
                self.consume_token();

                // Consume until newline
                self.consume_until_after(Token::Newline);
            }
        }

        self.builder.finish_node(); // IfStatement
    }

    /// Parse an `ElseIf` clause: `ElseIf` condition `Then` ...
    pub(super) fn parse_elseif_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseIfClause.to_raw());

        // Consume `ElseIf` keyword
        self.consume_token();

        // Consume any whitespace after `ElseIf`
        self.consume_whitespace();

        // Parse the conditional expression
        self.parse_expression();

        // Consume whitespace before Then
        self.consume_whitespace();

        // Consume `Then` if present
        if self.at_token(Token::ThenKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Then
        self.consume_whitespace();

        // Consume the newline after Then
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse body until `End If`, `Else`, or another `ElseIf`
        self.parse_statement_list(|parser| {
            parser.at_token(Token::ElseIfKeyword)
                || parser.at_token(Token::ElseKeyword)
                || (parser.at_token(Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(Token::IfKeyword))
        });

        self.builder.finish_node(); // ElseIfClause
    }

    /// Parse an `Else` clause: `Else` ...
    pub(super) fn parse_else_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseClause.to_raw());

        // Consume `Else` keyword
        self.consume_token();

        // Consume any whitespace after `Else`
        self.consume_whitespace();

        // Consume the newline after `Else`
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse body until `End If`
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::IfKeyword)
        });

        self.builder.finish_node(); // `ElseClause`
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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("ThenKeyword"));
    }

    #[test]
    fn inline_if_then_call() {
        let source = r"
Sub Test()
    If enabled Then Call DoSomething
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("AssignmentStatement"));
    }

    #[test]
    fn inline_if_then_set() {
        let source = r"
Sub Test()
    If obj Is Nothing Then Set obj = New MyClass
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ExitKeyword"));
    }

    #[test]
    fn inline_if_then_multiple_statements() {
        let source = r"
Sub Test()
    If condition Then x = 1: y = 2
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

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
        let source = r"
Sub Test()
    If x > 0 Then GoTo Label1
Label1:
    x = 1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("Whitespace"));
        assert!(debug.contains("Newline"));
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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn inline_if_then_call_with_args() {
        let source = r"
Sub Test()
    If ready Then Call Process(x, y, z)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("CallStatement"));
    }

    #[test]
    fn inline_if_then_nested_calls() {
        let source = r"
Sub Test()
    If value > 0 Then result = Calculate(value)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("AssignmentStatement"));
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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("GotoStatement"));
    }

    #[test]
    fn inline_if_not_condition() {
        let source = r"
Sub Test()
    If Not IsValid Then Exit Sub
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ExitKeyword"));
    }

    #[test]
    fn binary_conditional() {
        let source = r"Sub Test()
    If x = 5 Then
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind() == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a StatementList
        let parse_statement_list = sub_statement
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::StatementList)
            .expect("SubStatement should contain a StatementList");

        // The StatementList should contain an IfStatement
        let if_statement = parse_statement_list
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::IfStatement)
            .expect("StatementList should contain an IfStatement");

        // The IfStatement should contain a BinaryExpression
        let binary_conditional = if_statement
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::BinaryExpression)
            .expect("IfStatement should contain a BinaryExpression");

        // Verify the BinaryExpression structure
        assert_eq!(binary_conditional.kind(), SyntaxKind::BinaryExpression);
        assert!(
            !binary_conditional.is_token(),
            "BinaryExpression should be a node, not a token"
        );

        // Verify the BinaryExpression contains the expected elements:
        // whitespace, identifier "x", whitespace, "=", whitespace, number "5", whitespace
        assert!(binary_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IdentifierExpression && c.text() == "x"));
        assert!(binary_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::EqualityOperator));

        let literal_expr = binary_conditional
            .children()
            .iter()
            .find(|c| c.kind() == SyntaxKind::NumericLiteralExpression)
            .expect("BinaryExpression should contain a NumericLiteralExpression");

        assert!(literal_expr
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IntegerLiteral && c.text() == "5"));
    }

    #[test]
    fn unary_conditional() {
        let source = r"Sub Test()
    If Not isEmpty(x) Then
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind() == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a StatementList
        let statement_list = sub_statement
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::StatementList)
            .expect("SubStatement should contain a StatementList");

        // The StatementList should contain an IfStatement
        let if_statement = statement_list
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::IfStatement)
            .expect("StatementList should contain an IfStatement");

        // The IfStatement should contain a UnaryExpression
        let unary_conditional = if_statement
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::UnaryExpression)
            .expect("IfStatement should contain a UnaryExpression");

        // Verify the UnaryExpression structure
        assert_eq!(unary_conditional.kind(), SyntaxKind::UnaryExpression);
        assert!(
            !unary_conditional.is_token(),
            "UnaryExpression should be a node, not a token"
        );

        // Verify the UnaryExpression contains the expected elements:
        // whitespace, Not keyword, whitespace, CallExpression
        assert!(unary_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::NotKeyword));

        let call_expr = unary_conditional
            .children()
            .iter()
            .find(|c| c.kind() == SyntaxKind::CallExpression)
            .expect("UnaryExpression should contain a CallExpression");

        assert!(call_expr
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::Identifier && c.text() == "isEmpty"));
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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Navigate the tree structure
        let children = cst.children();

        // Find the SubStatement node
        let sub_statement = children
            .iter()
            .find(|child| child.kind() == SyntaxKind::SubStatement)
            .expect("Should have a SubStatement node");

        // The SubStatement should contain a StatementList
        let statement_list = sub_statement
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::StatementList)
            .expect("SubStatement should contain a StatementList");

        // Find the outer IfStatement in the StatementList
        let outer_if = statement_list
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::IfStatement)
            .expect("StatementList should contain an outer IfStatement");

        // Verify outer If has a BinaryExpression (x > 0)
        let outer_conditional = outer_if
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::BinaryExpression)
            .expect("Outer IfStatement should contain a BinaryExpression");
        assert!(outer_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IdentifierExpression && c.text() == "x"));
        assert!(outer_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::GreaterThanOperator));

        // Find the StatementList inside the outer If
        let outer_statement_list = outer_if
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::StatementList)
            .expect("Outer IfStatement should contain a StatementList");

        // Find the inner IfStatement (nested within the outer If's StatementList)
        let inner_if = outer_statement_list
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::IfStatement)
            .expect("Outer StatementList should contain a nested IfStatement");
        // Verify inner If has a BinaryExpression (y > 0)
        let inner_conditional = inner_if
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::BinaryExpression)
            .expect("Inner IfStatement should contain a BinaryExpression");
        assert!(inner_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IdentifierExpression && c.text() == "y"));
        assert!(inner_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::GreaterThanOperator));

        // Verify inner If has ElseIf clause
        let inner_elseif = inner_if
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::ElseIfClause)
            .expect("Inner IfStatement should contain an ElseIfClause");

        // Verify inner ElseIf has a BinaryExpression (y < 0)
        let inner_elseif_conditional = inner_elseif
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::BinaryExpression)
            .expect("Inner ElseIfClause should contain a BinaryExpression");
        assert!(inner_elseif_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IdentifierExpression && c.text() == "y"));
        assert!(inner_elseif_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::LessThanOperator));

        // Verify inner If has Else clause
        assert!(
            inner_if
                .children()
                .iter()
                .any(|child| child.kind() == SyntaxKind::ElseClause),
            "Inner IfStatement should contain an ElseClause"
        );

        // Verify outer If has ElseIf clause
        let outer_elseif = outer_if
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::ElseIfClause)
            .expect("Outer IfStatement should contain an ElseIfClause");

        // Verify outer ElseIf has a BinaryExpression (x < 0)
        let outer_elseif_conditional = outer_elseif
            .children()
            .iter()
            .find(|child| child.kind() == SyntaxKind::BinaryExpression)
            .expect("Outer ElseIfClause should contain a BinaryExpression");

        assert!(outer_elseif_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::IdentifierExpression && c.text() == "x"));
        assert!(outer_elseif_conditional
            .children()
            .iter()
            .any(|c| c.kind() == SyntaxKind::LessThanOperator));

        // Verify outer If has Else clause
        assert!(
            outer_if
                .children()
                .iter()
                .any(|child| child.kind() == SyntaxKind::ElseClause),
            "Outer IfStatement should contain an ElseClause"
        );
    }
}
