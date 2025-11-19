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
//! Function statements are handled in the function_statements module.
//! Dim/ReDim and general Variable declarations are handled in the array_statements module.
//! Property statements are handled in the property_statements module.
//! Parameter lists are handled in the parameters module.
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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
    /// | Public   	  | Optional | Indicates that the Sub procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the procedure is not available outside the project. |
    /// | Private  	  | Optional | Indicates that the Sub procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend 	  | Optional | Used only in a class module. Indicates that the Sub procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static 	  | Optional | Indicates that the Sub procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Sub, even if they are used in the procedure. |
    /// | name 	      | Required | Name of the Sub; follows standard variable naming conventions. |
    /// | arglist 	  | Optional | List of variables representing arguments that are passed to the Sub procedure when it is called. Multiple variables are separated by commas. |
    /// | statements  | Optional | Any group of statements to be executed within the Sub procedure.
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
    pub(super) fn parse_sub_statement(&mut self) {
        // if we are now parsing a sub statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::SubStatement.to_raw());

        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword)
            || self.at_token(VB6Token::PrivateKeyword)
            || self.at_token(VB6Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Sub" keyword
        self.consume_token();

        // Consume any whitespace after "Sub"
        self.consume_whitespace();

        // Consume procedure name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End Sub"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::SubKeyword)
        });

        // Consume "End Sub" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Sub"
            self.consume_whitespace();

            // Consume "Sub"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
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
    fn cst_public_static_sub() {
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
    fn cst_public_sub() {
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
    fn cst_private_sub() {
        // Test Private Sub
        let source = "Private Sub Initialize()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::SubStatement);
        }
        assert!(cst.text().contains("Private Sub Initialize"));
    }
}
