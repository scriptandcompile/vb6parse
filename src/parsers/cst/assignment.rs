//! Assignment statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 assignment statements:
//! - Let statement: `Let x = 5` (optional keyword)
//! - Simple variable assignment: `x = 5`
//! - Property assignment: `obj.property = value`
//! - Array assignment: `arr(index) = value`

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an assignment statement.
    ///
    /// VB6 assignment statement syntax:
    /// - variableName = expression
    /// - object.property = expression
    /// - array(index) = expression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assignment-operator)
    pub(super) fn parse_assignment_statement(&mut self) {
        // Assignments can appear in both header and body, so we do not modify parsing_header here.

        self.builder
            .start_node(SyntaxKind::AssignmentStatement.to_raw());

        // Track if we're at the start or after a period (where keywords can be identifiers)
        let mut at_identifier_position = true;
        let mut last_was_period = false;

        // Consume everything until newline or colon (for inline If statements)
        // This includes: variable/property, "=", expression
        while !self.is_at_end()
            && !self.at_token(VB6Token::Newline)
            && !self.at_token(VB6Token::ColonOperator)
        {
            // In VB6, keywords can be used as identifiers in certain positions:
            // - At the start of an assignment (variable name)
            // - After a period (property/method name)
            if (at_identifier_position || last_was_period) && self.at_keyword() {
                self.consume_token_as_identifier();
                at_identifier_position = false;
                last_was_period = false;
            } else {
                // Check if this is a period
                last_was_period = self.at_token(VB6Token::PeriodOperator);
                
                // After whitespace, we're still in an identifier position
                if !self.at_token(VB6Token::Whitespace) {
                    at_identifier_position = false;
                }
                
                self.consume_token();
            }
        }

        // Consume the newline if present (but not colon - that's handled by caller)
        if self.at_token(VB6Token::Newline) {
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

        // Consume everything until newline or colon (for inline If statements)
        // This includes: variable/property, "=", expression
        while !self.is_at_end()
            && !self.at_token(VB6Token::Newline)
            && !self.at_token(VB6Token::ColonOperator)
        {
            self.consume_token();
        }

        // Consume the newline if present (but not colon - that's handled by caller)
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LetStatement
    }

    /// Check if the current position is at the start of an assignment statement.
    /// This looks ahead to see if there's an `=` operator (not part of a comparison).
    /// Note: Let statements are handled separately and should be checked first.
    pub(super) fn is_at_assignment(&self) -> bool {
        // Let statements are handled separately
        if self.at_token(VB6Token::LetKeyword) {
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
                VB6Token::Newline | VB6Token::EndOfLineComment | VB6Token::RemComment => {
                    // Reached end of line without finding assignment
                    return false;
                }
                VB6Token::EqualityOperator => {
                    // Found an = operator - this is likely an assignment
                    return true;
                }
                VB6Token::PeriodOperator => {
                    last_was_period = true;
                    at_start = false;
                    continue;
                }
                // Skip tokens that could appear in the left-hand side of an assignment
                VB6Token::Whitespace => {
                    continue;
                }
                VB6Token::Identifier
                | VB6Token::LeftParenthesis
                | VB6Token::RightParenthesis
                | VB6Token::Number
                | VB6Token::Comma => {
                    last_was_period = false;
                    at_start = false;
                    continue;
                }
                // After a period, keywords can be property names, so skip them
                _ if last_was_period => {
                    last_was_period = false;
                    at_start = false;
                    continue;
                }
                // At the start of a statement, keywords can be used as variable names
                _ if at_start && token.is_keyword() => {
                    at_start = false;
                    continue;
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
    fn test_simple_assignment() {
        let source = r#"
x = 5
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "x");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[4].text, "5");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_string_assignment() {
        let source = r#"
myName = "John"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "myName");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::StringLiteral
        );
        assert_eq!(cst.children()[1].children[4].text, "\"John\"");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_property_assignment() {
        let source = r#"
obj.subProperty = value
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // The assignment contains: obj.subProperty = value
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "obj");
        assert_eq!(
            cst.children()[1].children[1].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[2].text, "subProperty");
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[6].text, "value");
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_array_assignment() {
        let source = r#"
arr(0) = 100
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "arr");
        assert_eq!(
            cst.children()[1].children[1].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[2].text, "0");
        assert_eq!(
            cst.children()[1].children[3].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[7].text, "100");
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_multidimensional_array_assignment() {
        let source = r#"
matrix(i, j) = value
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "matrix");
        assert_eq!(
            cst.children()[1].children[1].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[2].text, "i");
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Comma);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[5].text, "j");
        assert_eq!(
            cst.children()[1].children[6].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[8].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[10].text, "value");
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_function_call() {
        let source = r#"
result = MyFunction(arg1, arg2)
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "result");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "MyFunction");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[6].text, "arg1");
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Comma);
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[9].text, "arg2");
        assert_eq!(
            cst.children()[1].children[10].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_expression() {
        let source = r#"
sum = a + b * c
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "sum");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "a");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[6].kind,
            SyntaxKind::AdditionOperator
        );
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[8].text, "b");
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[10].kind,
            SyntaxKind::MultiplicationOperator
        );
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[12].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[12].text, "c");
        assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_method_call() {
        let source = r#"
text = obj.GetText()
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "text" is converted to Identifier even though it's TextKeyword in the tokenizer
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "text");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "obj");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[6].text, "GetText");
        assert_eq!(
            cst.children()[1].children[7].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(
            cst.children()[1].children[8].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_nested_property() {
        let source = r#"
value = obj.SubObj.SubProperty
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "value");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "obj");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[6].text, "SubObj");
        assert_eq!(
            cst.children()[1].children[7].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[8].text, "SubProperty");
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_multiple_assignments() {
        let source = r#"
x = 1
y = 2
z = 3
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "x");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[4].text, "1");

        assert_eq!(cst.children()[2].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[2].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[2].children[0].text, "y");
        assert_eq!(cst.children()[2].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[2].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[2].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[2].children[4].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[2].children[4].text, "2");

        assert_eq!(cst.children()[3].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[3].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[3].children[0].text, "z");
        assert_eq!(cst.children()[3].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[3].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[3].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[3].children[4].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[3].children[4].text, "3");

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_preserves_whitespace() {
        let source = "x   =   5";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[0].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[0].children[0].text, "x");
        assert_eq!(cst.children()[0].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[0].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[0].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[0].children[4].kind, SyntaxKind::Number);

        // Verify whitespace is preserved
        assert_eq!(cst.text(), source);
    }

    #[test]
    fn test_assignment_in_function() {
        let source = r#"
Public Function Calculate()
    result = 42
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::FunctionStatement);
        assert_eq!(
            cst.children()[1].children[0].kind,
            SyntaxKind::PublicKeyword
        );
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::FunctionKeyword
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "Calculate");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::ParameterList
        );
        assert_eq!(
            cst.children()[1].children[5].children[0].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(
            cst.children()[1].children[5].children[1].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Newline);

        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::CodeBlock);
        assert_eq!(
            cst.children()[1].children[7].children[0].kind,
            SyntaxKind::Whitespace
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].kind,
            SyntaxKind::AssignmentStatement
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[0].kind,
            SyntaxKind::Identifier
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[0].text,
            "result"
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[1].kind,
            SyntaxKind::Whitespace
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[3].kind,
            SyntaxKind::Whitespace
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[4].kind,
            SyntaxKind::Number
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[4].text,
            "42"
        );
        assert_eq!(
            cst.children()[1].children[7].children[1].children[5].kind,
            SyntaxKind::Newline
        );

        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::EndKeyword);
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[10].kind,
            SyntaxKind::FunctionKeyword
        );
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_collection_access() {
        let source = r#"
item = Collection("Key")
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "item");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "Collection");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(
            cst.children()[1].children[6].kind,
            SyntaxKind::StringLiteral
        );
        assert_eq!(cst.children()[1].children[6].text, "\"Key\"");
        assert_eq!(
            cst.children()[1].children[7].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_dollar_sign_function() {
        let source = r#"
path = Environ$("TEMP")
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "path");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "Environ");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::DollarSign);
        assert_eq!(
            cst.children()[1].children[6].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(
            cst.children()[1].children[7].kind,
            SyntaxKind::StringLiteral
        );
        assert_eq!(cst.children()[1].children[7].text, "\"TEMP\"");
        assert_eq!(
            cst.children()[1].children[8].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_at_module_level() {
        let source = r#"
Option Explicit
x = 5
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert!(cst.children()[0].kind == SyntaxKind::Newline);

        assert!(cst.children()[1].kind == SyntaxKind::OptionStatement);
        assert!(cst.children()[1].children[0].kind == SyntaxKind::OptionKeyword);
        assert!(cst.children()[1].children[1].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[1].children[2].kind == SyntaxKind::ExplicitKeyword);
        assert!(cst.children()[1].children[3].kind == SyntaxKind::Newline);

        assert!(cst.children()[2].kind == SyntaxKind::AssignmentStatement);
        assert!(cst.children()[2].children[0].kind == SyntaxKind::Identifier);
        assert!(cst.children()[2].children[0].text == "x");
        assert!(cst.children()[2].children[1].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[2].children[2].kind == SyntaxKind::EqualityOperator);
        assert!(cst.children()[2].children[3].kind == SyntaxKind::Whitespace);
        assert!(cst.children()[2].children[4].kind == SyntaxKind::Number);
        assert!(cst.children()[2].children[4].text == "5");
        assert!(cst.children()[2].children[5].kind == SyntaxKind::Newline);

        // Verify the parsed tree can be converted back to the original source
        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_numeric_literal() {
        let source = r#"
pi = 3.14159
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "pi");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[4].text, "3");
        assert_eq!(
            cst.children()[1].children[5].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[6].text, "14159");
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_concatenation() {
        let source = r#"
fullName = firstName & " " & lastName
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "fullName");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[4].text, "firstName");
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Ampersand);
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[8].kind,
            SyntaxKind::StringLiteral
        );
        assert_eq!(cst.children()[1].children[8].text, "\" \"");
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::Ampersand);
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[12].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[12].text, "lastName");
        assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_to_type_member() {
        let source = r#"
person.Age = 25
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "person");
        assert_eq!(
            cst.children()[1].children[1].kind,
            SyntaxKind::PeriodOperator
        );
        assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[2].text, "Age");
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Number);
        assert_eq!(cst.children()[1].children[6].text, "25");
        assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_assignment_with_parenthesized_expression() {
        let source = r#"
result = (a + b) * c
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "result");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[2].kind,
            SyntaxKind::EqualityOperator
        );
        assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[4].kind,
            SyntaxKind::LeftParenthesis
        );
        assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[5].text, "a");
        assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[7].kind,
            SyntaxKind::AdditionOperator
        );
        assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[9].text, "b");
        assert_eq!(
            cst.children()[1].children[10].kind,
            SyntaxKind::RightParenthesis
        );
        assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
        assert_eq!(
            cst.children()[1].children[12].kind,
            SyntaxKind::MultiplicationOperator
        );
        assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Whitespace);
        assert_eq!(cst.children()[1].children[14].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[14].text, "c");
        assert_eq!(cst.children()[1].children[15].kind, SyntaxKind::Newline);

        assert_eq!(cst.text().trim(), source.trim());
    }

    #[test]
    fn test_let_simple() {
        let source = r#"
Sub Test()
    Let x = 5
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("LetKeyword"));
        assert!(debug.contains("x"));
    }

    #[test]
    fn test_let_module_level() {
        let source = "Let myVar = 10\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("myVar"));
    }

    #[test]
    fn test_let_string() {
        let source = r#"
Sub Test()
    Let myName = "John"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("myName"));
        assert!(debug.contains("John"));
    }

    #[test]
    fn test_let_expression() {
        let source = r#"
Sub Test()
    Let result = x + y * 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn test_let_property_access() {
        let source = r#"
Sub Test()
    Let obj.Value = 100
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("obj"));
        assert!(debug.contains("Value"));
    }

    #[test]
    fn test_let_array_element() {
        let source = r#"
Sub Test()
    Let arr(5) = 42
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("arr"));
    }

    #[test]
    fn test_let_preserves_whitespace() {
        let source = "    Let    x    =    5    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Let    x    =    5    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn test_let_with_comment() {
        let source = r#"
Sub Test()
    Let counter = 0 ' Initialize counter
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn test_let_in_if_statement() {
        let source = r#"
Sub Test()
    If x > 0 Then
        Let y = x
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn test_let_inline_if() {
        let source = r#"
Sub Test()
    If condition Then Let x = 5
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
    }

    #[test]
    fn test_multiple_let_statements() {
        let source = r#"
Sub Test()
    Let a = 1
    Let b = 2
    Let c = 3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let let_count = debug.matches("LetStatement").count();
        assert_eq!(let_count, 3);
    }

    #[test]
    fn test_let_with_function_call() {
        let source = r#"
Sub Test()
    Let result = Calculate(x, y)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("result"));
        assert!(debug.contains("Calculate"));
    }

    #[test]
    fn test_let_with_concatenation() {
        let source = r#"
Sub Test()
    Let fullName = firstName & " " & lastName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LetStatement"));
        assert!(debug.contains("fullName"));
        assert!(debug.contains("firstName"));
        assert!(debug.contains("lastName"));
    }

    #[test]
    fn test_keyword_as_variable_name() {
        let source = r#"
text = "hello"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "text" keyword should be converted to Identifier
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "text");
    }

    #[test]
    fn test_keyword_as_property_name() {
        let source = r#"
obj.text = "hello"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "obj");
        assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::PeriodOperator);
        // "text" keyword after period should be converted to Identifier
        assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[2].text, "text");
    }

    #[test]
    fn test_database_keyword_as_variable() {
        let source = r#"
database = "mydb.mdb"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
        assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
        // "database" keyword should be converted to Identifier
        assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
        assert_eq!(cst.children()[1].children[0].text, "database");
    }
}
