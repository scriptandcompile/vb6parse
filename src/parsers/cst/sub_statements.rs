//! Sub statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Sub statements with syntax:
//!
//! \[ Public | Private | Friend \] \[ Static \] Sub name \[ ( arglist ) \]
//! \[ statements \]
//! \[ Exit Sub \]
//! \[ statements \]
//! End Sub
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 subroutine with syntax:
    ///
    /// \[ Public | Private | Friend \] \[ Static \] Sub name \[ ( arglist ) \]
    /// \[ statements \]
    /// \[ Exit Sub \]
    /// \[ statements \]
    /// End Sub
    ///
    /// The Sub statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the Sub procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the procedure is not available outside the project. |
    /// | Private     | Optional | Indicates that the Sub procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend      | Optional | Used only in a class module. Indicates that the Sub procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static      | Optional | Indicates that the Sub procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Sub, even if they are used in the procedure. |
    /// | name        | Required | Name of the Sub; follows standard variable naming conventions. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the Sub procedure when it is called. Multiple variables are separated by commas. |
    /// | statements  | Optional | Any group of statements to be executed within the Sub procedure.
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ `ByVal` | `ByRef` \] \[ `ParamArray` \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
    pub(super) fn parse_sub_statement(&mut self) {
        // if we are now parsing a sub statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::SubStatement.to_raw());

        // Consume optional Public/Private/Friend keyword
        if self.at_token(Token::PublicKeyword)
            || self.at_token(Token::PrivateKeyword)
            || self.at_token(Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Sub" keyword
        self.consume_token();

        // Consume any whitespace after "Sub"
        self.consume_whitespace();

        // Consume procedure name (keywords can be used as procedure names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        // Parse body until "End Sub"
        self.parse_code_block(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::SubKeyword)
        });

        // Consume "End Sub" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Sub"
            self.consume_whitespace();

            // Consume "Sub"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // SubStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn sub_all_modifier_combinations() {
        // Test all valid sub modifier combinations
        let test_cases = vec![
            ("Public Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
            ("Private Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
            ("Friend Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
            ("Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
            (
                "Public Static Sub Test()\nEnd Sub\n",
                SyntaxKind::SubStatement,
            ),
            (
                "Private Static Sub Test()\nEnd Sub\n",
                SyntaxKind::SubStatement,
            ),
            (
                "Friend Static Sub Test()\nEnd Sub\n",
                SyntaxKind::SubStatement,
            ),
        ];

        for (source, expected_kind) in test_cases {
            let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

            assert_eq!(cst.child_count(), 1, "Code: {}", source);
            if let Some(child) = cst.child_at(0) {
                assert_eq!(child.kind, expected_kind, "Code: {}", source);
            }
        }
    }

    #[test]
    fn sub_public_static() {
        // Test Public Static Sub
        let source = "Public Static Sub Initialize()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Public Static Sub Initialize"));
    }

    #[test]
    fn sub_public() {
        // Test Public Sub
        let source = "Public Sub Initialize()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Public Sub Initialize"));
    }

    #[test]
    fn sub_private() {
        // Test Private Sub
        let source = "Private Sub Initialize()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Private Sub Initialize"));
    }

    #[test]
    fn sub_simple_no_params() {
        // Test simple sub with no parameters
        let source = "Sub DoSomething()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
    }

    #[test]
    fn sub_with_params() {
        // Test sub with parameters
        let source = "Sub SetValue(ByVal x As Integer, ByVal y As Integer)\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("ByVal x As Integer"));
    }

    #[test]
    fn sub_with_exit_sub() {
        // Test sub with Exit Sub statement
        let source = "Sub Validate(x As Integer)\n    If x < 0 Then\n        Exit Sub\n    End If\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Exit Sub"));
    }

    #[test]
    fn sub_friend_modifier() {
        // Test Friend Sub
        let source = "Friend Sub ProcessData()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Friend Sub"));
    }

    #[test]
    fn sub_static_modifier() {
        // Test Static Sub
        let source = "Static Sub Counter()\n    Dim count As Integer\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Static Sub"));
    }

    #[test]
    fn sub_with_body() {
        // Test sub with body statements
        let source = "Sub Calculate()\n    Dim x As Integer\n    x = 10\n    MsgBox x\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Dim x As Integer"));
        assert!(cst.text().contains("x = 10"));
    }

    #[test]
    fn sub_with_optional_params() {
        // Test sub with optional parameters
        let source = "Sub Process(ByVal x As Integer, Optional ByVal y As Integer = 0)\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Optional"));
    }

    #[test]
    fn sub_with_keyword_as_name() {
        let source = r#"Sub Text()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SubStatement"));
        // "Text" keyword should be converted to Identifier when used as procedure name
        assert!(debug.contains("Identifier@4..8 \"Text\""));
    }
}
