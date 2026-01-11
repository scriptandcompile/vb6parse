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

        // Consume any leading whitespace
        self.consume_whitespace();

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

        // Consume any leading whitespace
        self.consume_whitespace();

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
mod tests {
    use crate::*;

    #[test]
    fn simple_assignment() {
        let source = r"
x = 5
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn string_assignment() {
        let source = r#"
myName = "John"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn property_assignment() {
        let source = r"
obj.subProperty = value
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn array_assignment() {
        let source = r"
arr(0) = 100
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multidimensional_array_assignment() {
        let source = r"
matrix(i, j) = value
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_function_call() {
        let source = r"
result = MyFunction(arg1, arg2)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_expression() {
        let source = r"
sum = a + b * c
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_method_call() {
        let source = r"
text = obj.GetText()
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_nested_property() {
        let source = r"
value = obj.SubObj.SubProperty
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_assignments() {
        let source = r"
x = 1
y = 2
z = 3
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_preserves_whitespace() {
        let source = "x   =   5";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_in_function() {
        let source = r"
Public Function Calculate()
    result = 42
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_collection_access() {
        let source = r#"
item = Collection("Key")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_dollar_sign_function() {
        let source = r#"
path = Environ$("TEMP")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_at_module_level() {
        let source = r"
Option Explicit
x = 5
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_numeric_literal() {
        let source = r"
pi = 3.14159
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_concatenation() {
        let source = r#"
fullName = firstName & " " & lastName
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_to_type_member() {
        let source = r"
person.Age = 25
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn assignment_with_parenthesized_expression() {
        let source = r"
result = (a + b) * c
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_simple() {
        let source = r"
Sub Test()
    Let x = 5
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_module_level() {
        let source = "Let myVar = 10\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_string() {
        let source = r#"
Sub Test()
    Let myName = "John"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_expression() {
        let source = r"
Sub Test()
    Let result = x + y * 2
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_property_access() {
        let source = r"
Sub Test()
    Let obj.Value = 100
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_array_element() {
        let source = r"
Sub Test()
    Let arr(5) = 42
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_preserves_whitespace() {
        let source = "    Let    x    =    5    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_with_comment() {
        let source = r"
Sub Test()
    Let counter = 0 ' Initialize counter
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_inline_if() {
        let source = r"
Sub Test()
    If condition Then Let x = 5
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_with_function_call() {
        let source = r"
Sub Test()
    Let result = Calculate(x, y)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn let_with_concatenation() {
        let source = r#"
Sub Test()
    Let fullName = firstName & " " & lastName
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn keyword_as_variable_name() {
        let source = r#"
text = "hello"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn keyword_as_property_name() {
        let source = r#"
obj.text = "hello"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn database_keyword_as_variable() {
        let source = r#"
database = "mydb.mdb"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/assignment");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
