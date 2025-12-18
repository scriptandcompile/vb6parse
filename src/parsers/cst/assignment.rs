//! Assignment statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 assignment statements:
//! - Let statement: `Let x = 5` (optional keyword)
//! - Simple variable assignment: `x = 5`
//! - Property assignment: `obj.property = value`
//! - Array assignment: `arr(index) = value`

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse an assignment statement.
    ///
    /// VB6 assignment statement syntax:
    /// - variableName = expression
    /// - object.property = expression
    /// - array(index) = expression
    ///
    pub(super) fn parse_assignment_statement(&mut self) {
        // Assignments can appear in both header and body, so we do not modify parsing_header here.

        self.builder
            .start_node(SyntaxKind::AssignmentStatement.to_raw());

        // Parse left-hand side - use parse_lvalue which stops before =
        self.parse_lvalue();

        // Skip whitespace
        self.consume_whitespace();

        // Consume the equals sign
        if self.at_token(Token::EqualityOperator) {
            self.consume_token();
        }

        // Skip whitespace after =
        self.consume_whitespace();

        // Parse right-hand side (value expression)
        self.parse_expression();

        // Consume the newline if present (but not colon - that's handled by caller)
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AssignmentStatement
    }

    /// Parse a Let statement.
    ///
    /// VB6 Let statement syntax:
    /// - Let variableName = expression
    ///
    /// The Let keyword is optional in VB6 and is provided for backward compatibility.
    /// Most modern VB6 code omits the Let keyword.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/let-statement)
    pub(super) fn parse_let_statement(&mut self) {
        // Let statements can appear in both header and body, so we do not modify parsing_header here.

        self.builder.start_node(SyntaxKind::LetStatement.to_raw());

        // Consume "Let" keyword
        self.consume_token();

        // Parse left-hand side
        self.parse_lvalue();

        // Skip whitespace
        self.consume_whitespace();

        // Consume "="
        if self.at_token(Token::EqualityOperator) {
            self.consume_token();
        }

        // Skip whitespace
        self.consume_whitespace();

        // Parse right-hand side
        self.parse_expression();

        // Consume the newline if present (but not colon - that's handled by caller)
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LetStatement
    }

    /// Check if the current position is at the start of an assignment statement.
    /// This looks ahead to see if there's an `=` operator (not part of a comparison).
    /// Note: Let statements are handled separately and should be checked first.
    pub(super) fn is_at_assignment(&self) -> bool {
        // Let statements are handled separately
        if self.at_token(Token::LetKeyword) {
            return false;
        }

        // Look ahead through the tokens to find an = operator before a newline
        // We need to skip: identifiers, periods, parentheses, array indices, etc.
        // Note: In VB6, keywords can be used as property/member names (e.g., obj.Property = value)
        // and also as variable names (e.g., text = "hello")
        let mut last_was_period = false;
        let mut at_start = true;

        for (_text, token) in self.tokens.iter().skip(self.pos) {
            match token {
                Token::Newline | Token::EndOfLineComment | Token::RemComment => {
                    // Reached end of line without finding assignment
                    return false;
                }
                Token::EqualityOperator => {
                    // Found an = operator - this is likely an assignment
                    return true;
                }
                Token::PeriodOperator => {
                    last_was_period = true;
                    at_start = false;
                }
                // Skip tokens that could appear in the left-hand side of an assignment
                Token::Whitespace => {}
                Token::Identifier
                | Token::LeftParenthesis
                | Token::RightParenthesis
                | Token::IntegerLiteral
                | Token::LongLiteral
                | Token::SingleLiteral
                | Token::DoubleLiteral
                | Token::Comma => {
                    last_was_period = false;
                    at_start = false;
                }
                // After a period, keywords can be property names, so skip them
                _ if last_was_period => {
                    last_was_period = false;
                    at_start = false;
                }
                // At the start of a statement, keywords can be used as variable names
                _ if at_start && token.is_keyword() => {
                    at_start = false;
                }
                // If we hit other operators, it's not an assignment
                _ => {
                    return false;
                }
            }
        }
        false
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::cst::ConcreteSyntaxTree;
    use crate::parsers::SyntaxKind;

    #[test]
    fn simple_assignment() {
        let source = r"
x = 5
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // Left side: IdentifierExpression containing Identifier
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert_eq!(
            cst.children()[1].children[0].children[0].kind,
            SyntaxKind::Identifier
        );
        assert_eq!(cst.children()[1].children[0].children[0].text, "x");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: NumericLiteralExpression containing IntegerLiteral
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );
        assert_eq!(
            cst.children()[1].children[4].children[0].kind,
            SyntaxKind::IntegerLiteral
        );
        assert_eq!(cst.children()[1].children[4].children[0].text, "5");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn string_assignment() {
        let source = r#"
myName = "John"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert_eq!(
            cst.children()[1].children[0].children[0].kind,
            SyntaxKind::Identifier
        );
        assert_eq!(cst.children()[1].children[0].children[0].text, "myName");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::StringLiteralExpression
        );
        assert_eq!(
            cst.children()[1].children[4].children[0].kind,
            SyntaxKind::StringLiteral
        );
        assert_eq!(cst.children()[1].children[4].children[0].text, "\"John\"");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn property_assignment() {
        let source = r"
obj.subProperty = value
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // The assignment contains: MemberAccessExpression = IdentifierExpression
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::MemberAccessExpression
        );
        assert!(cst.children()[1].children[0].text.contains("obj"));
        assert!(cst.children()[1].children[0].text.contains("subProperty"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::IdentifierExpression
        );
        assert_eq!(
            cst.children()[1].children[4].children[0].kind,
            SyntaxKind::Identifier
        );
        assert_eq!(cst.children()[1].children[4].children[0].text, "value");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn array_assignment() {
        let source = r"
arr(0) = 100
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // Left side: CallExpression for array access
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[0].text.contains("arr"));
        assert!(cst.children()[1].children[0].text.contains('0'));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: NumericLiteralExpression
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );
        assert!(cst.children()[1].children[4].text.contains("100"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn multidimensional_array_assignment() {
        let source = r"
matrix(i, j) = value
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[0].text.contains("matrix"));
        assert!(cst.children()[1].children[0].text.contains('i'));
        assert!(cst.children()[1].children[0].text.contains('j'));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[4].text.contains("value"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_function_call() {
        let source = r"
result = MyFunction(arg1, arg2)
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("result"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[4].text.contains("MyFunction"));
        assert!(cst.children()[1].children[4].text.contains("arg1"));
        assert!(cst.children()[1].children[4].text.contains("arg2"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_expression() {
        let source = r"
sum = a + b * c
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("sum"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: BinaryExpression (a + (b * c))
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::BinaryExpression
        );
        assert!(cst.children()[1].children[4].text.contains('a'));
        assert!(cst.children()[1].children[4].text.contains('+'));
        assert!(cst.children()[1].children[4].text.contains('b'));
        assert!(cst.children()[1].children[4].text.contains('*'));
        assert!(cst.children()[1].children[4].text.contains('c'));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_method_call() {
        let source = r"
text = obj.GetText()
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "text" is converted to IdentifierExpression
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("text"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: CallExpression wrapping MemberAccessExpression
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[4].text.contains("obj"));
        assert!(cst.children()[1].children[4].text.contains("GetText"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_nested_property() {
        let source = r"
value = obj.SubObj.SubProperty
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("value"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: nested MemberAccessExpression
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::MemberAccessExpression
        );
        assert!(cst.children()[1].children[4].text.contains("obj"));
        assert!(cst.children()[1].children[4].text.contains("SubObj"));
        assert!(cst.children()[1].children[4].text.contains("SubProperty"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn multiple_assignments() {
        let source = r"
x = 1
y = 2
z = 3
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains('x'));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );

        assert_eq!(cst.children()[2].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[2].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[2].children[0].text.contains('y'));
        assert_eq!(cst.children()[2].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[2].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[2].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[2].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );

        assert_eq!(cst.children()[3].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[3].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[3].children[0].text.contains('z'));
        assert_eq!(cst.children()[3].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[3].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[3].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[3].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_preserves_whitespace() {
        let source = "x   =   5";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[0].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[0].children[0].text.contains('x'));
        assert_eq!(cst.children()[0].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[0].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[0].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[0].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );

        // Verify whitespace is preserved
        assert_eq!(cst.text(), source);
    }

    #[test]
    fn assignment_in_function() {
        let source = r"
Public Function Calculate()
    result = 42
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let assignment = &cst.children()[1].children[7].children[1];
        assert_eq!(assignment.kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            assignment.children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(assignment.children[0].text.contains("result"));
        assert_eq!(assignment.children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(assignment.children[2].kind, SyntaxKind::EqualityOperator);
        assert_eq!(assignment.children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            assignment.children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );
        assert!(assignment.children[4].text.contains("42"));
        assert_eq!(assignment.children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_collection_access() {
        let source = r#"
item = Collection("Key")
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("item"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[4].text.contains("Collection"));
        assert!(cst.children()[1].children[4].text.contains("Key"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_dollar_sign_function() {
        let source = r#"
path = Environ$("TEMP")
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("path"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::CallExpression
        );
        assert!(cst.children()[1].children[4].text.contains("Environ$"));
        assert!(cst.children()[1].children[4].text.contains("TEMP"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_at_module_level() {
        let source = r"
Option Explicit
x = 5
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert!(cst.children()[0].kind == SyntaxKind::Newline);

        assert!(cst.children()[1].kind == SyntaxKind::OptionStatement);
        assert!(cst.children()[1].children[0].kind == SyntaxKind::OptionKeyword);
        assert!(cst.children()[1].children[1].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[1].children[2].kind == SyntaxKind::ExplicitKeyword);
        assert!(cst.children()[1].children[3].kind == SyntaxKind::Newline);

        assert!(cst.children()[2].kind == SyntaxKind::AssignmentStatement);
        assert!(cst.children()[2].children[0].kind == SyntaxKind::IdentifierExpression);
        assert!(cst.children()[2].children[0].text.contains('x'));
        assert!(cst.children()[2].children[1].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[2].children[2].kind == SyntaxKind::EqualityOperator);
        assert!(cst.children()[2].children[3].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[2].children[4].kind == SyntaxKind::NumericLiteralExpression);
        assert!(cst.children()[2].children[4].text.contains('5'));
        assert!(cst.children()[2].children[5].kind == SyntaxKind::Newline);

        // Verify the parsed tree can be converted back to the original source
        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_numeric_literal() {
        let source = r"
pi = 3.14159
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("pi"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );
        assert!(cst.children()[1].children[4].text.contains("3.14159"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_concatenation() {
        let source = r#"
fullName = firstName & " " & lastName
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("fullName"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: BinaryExpression with concatenation
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::BinaryExpression
        );
        assert!(cst.children()[1].children[4].text.contains("firstName"));
        assert!(cst.children()[1].children[4].text.contains('&'));
        assert!(cst.children()[1].children[4].text.contains("lastName"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_to_type_member() {
        let source = r"
person.Age = 25
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::MemberAccessExpression
        );
        assert!(cst.children()[1].children[0].text.contains("person"));
        assert!(cst.children()[1].children[0].text.contains("Age"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::NumericLiteralExpression
        );
        assert!(cst.children()[1].children[4].text.contains("25"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn assignment_with_parenthesized_expression() {
        let source = r"
result = (a + b) * c
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("result"));
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        // Right side: BinaryExpression with ParenthesizedExpression
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::BinaryExpression
        );
        let debug = cst.debug_tree();
        assert!(debug.contains("ParenthesizedExpression"));
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn let_simple() {
        let source = r"
Sub Test()
    Let x = 5
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("LetKeyword"));
        assert!(debug.contains('x'));
    }

    #[test]
    fn let_module_level() {
        let source = "Let myVar = 10\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("myVar"));
    }

    #[test]
    fn let_string() {
        let source = r#"
Sub Test()
    Let myName = "John"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("myName"));
        assert!(debug.contains("John"));
    }

    #[test]
    fn let_expression() {
        let source = r"
Sub Test()
    Let result = x + y * 2
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn let_property_access() {
        let source = r"
Sub Test()
    Let obj.Value = 100
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("obj"));
        assert!(debug.contains("Value"));
    }

    #[test]
    fn let_array_element() {
        let source = r"
Sub Test()
    Let arr(5) = 42
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("arr"));
    }

    #[test]
    fn let_preserves_whitespace() {
        let source = "    Let    x    =    5    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Let    x    =    5    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn let_with_comment() {
        let source = r"
Sub Test()
    Let counter = 0 ' Initialize counter
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn let_in_if_statement() {
        let source = r"
Sub Test()
    If x > 0 Then
        Let y = x
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn let_inline_if() {
        let source = r"
Sub Test()
    If condition Then Let x = 5
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn multiple_let_statements() {
        let source = r"
Sub Test()
    Let a = 1
    Let b = 2
    Let c = 3
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let let_count = debug.matches("LetStatement").count();
        assert_eq!(let_count, 3);
    }

    #[test]
    fn let_with_function_call() {
        let source = r"
Sub Test()
    Let result = Calculate(x, y)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("result"));
        assert!(debug.contains("Calculate"));
    }

    #[test]
    fn let_with_concatenation() {
        let source = r#"
Sub Test()
    Let fullName = firstName & " " & lastName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("fullName"));
        assert!(debug.contains("firstName"));
        assert!(debug.contains("lastName"));
    }

    #[test]
    fn keyword_as_variable_name() {
        let source = r#"
text = "hello"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "text" keyword should be converted to IdentifierExpression
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("text"));
    }

    #[test]
    fn keyword_as_property_name() {
        let source = r#"
obj.text = "hello"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::MemberAccessExpression
        );
        assert!(cst.children()[1].children[0].text.contains("obj"));
        assert!(cst.children()[1].children[0].text.contains("text"));
    }

    #[test]
    fn database_keyword_as_variable() {
        let source = r#"
database = "mydb.mdb"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "database" keyword should be converted to IdentifierExpression
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::IdentifierExpression
        );
        assert!(cst.children()[1].children[0].text.contains("database"));
    }
}
