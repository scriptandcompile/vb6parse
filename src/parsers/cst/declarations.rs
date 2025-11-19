//! Declaration parsing for VB6 CST.
//!
//! This module handles parsing of VB6 declarations:
//! - Function statements
//! - Sub statements
//! - Property statements (Get, Let, Set)
//! - Parameter lists
//!
//! Dim/ReDim and general Variable declarations are handled in the array_statements module.

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

    /// Parse a Visual Basic 6 function with syntax:
    ///
    /// \[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// \[ Exit Function \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// End Function
    ///
    /// The Function statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public   	  | Optional | Indicates that the Function procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private, the procedure is not available outside the project. |
    /// | Private  	  | Optional | Indicates that the Function procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend 	  | Optional | Used only in a class module. Indicates that the Function procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static 	  | Optional | Indicates that the Function procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Function, even if they are used in the procedure. |
    /// | name 	      | Required | Name of the Function; follows standard variable naming conventions. |
    /// | arglist 	  | Optional | List of variables representing arguments that are passed to the Function procedure when it is called. Multiple variables are separated by commas. |
    /// | type 	      | Optional | Data type of the value returned by the Function procedure; may be Byte, Boolean, Integer, Long, Currency, Single, Double, Decimal (not currently supported), Date, String (except fixed length), Object, Variant, or any user-defined type. |
    /// | statements  | Optional | Any group of statements to be executed within the Function procedure.
    /// | expression  | Optional | Return value of the Function. |
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)
    pub(super) fn parse_function_statement(&mut self) {
        // if we are now parsing a function statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::FunctionStatement.to_raw());

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

        // Consume "Function" keyword
        self.consume_token();

        // Consume any whitespace after "Function"
        self.consume_whitespace();

        // Consume function name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (includes "As Type" if present)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End Function"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::FunctionKeyword)
        });

        // Consume "End Function" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Function"
            self.consume_whitespace();

            // Consume "Function"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // FunctionStatement
    }

    /// Parse a parameter list: (param1 As Type, param2 As Type)
    pub(super) fn parse_parameter_list(&mut self) {
        self.builder.start_node(SyntaxKind::ParameterList.to_raw());

        // Consume "("
        self.consume_token();

        // Consume everything until ")"
        let mut depth = 1;
        while !self.is_at_end() && depth > 0 {
            if self.at_token(VB6Token::LeftParenthesis) {
                depth += 1;
            } else if self.at_token(VB6Token::RightParenthesis) {
                depth -= 1;
            }

            self.consume_token();

            if depth == 0 {
                break;
            }
        }

        self.builder.finish_node(); // ParameterList
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn cst_distinguishes_declarations_from_functions() {
        // Test that Private declaration and Private Function are correctly distinguished
        let source =
            "Private myVar As Integer\nPrivate Function GetVar() As Integer\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 2);

        // First child should be a DimStatement (declaration)
        if let Some(first_child) = cst.child_at(0) {
            assert_eq!(first_child.kind, SyntaxKind::DimStatement);
        }

        // Second child should be a FunctionStatement
        if let Some(second_child) = cst.child_at(1) {
            assert_eq!(second_child.kind, SyntaxKind::FunctionStatement);
        }

        assert!(cst.text().contains("Private myVar As Integer"));
        assert!(cst.text().contains("Private Function GetVar"));
    }

    #[test]
    fn cst_all_function_modifier_combinations() {
        // Test all valid function/sub modifier combinations
        let test_cases = vec![
            (
                "Public Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Private Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Friend Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Static Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Public Static Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Private Static Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
            (
                "Friend Static Function Test() As Integer\nEnd Function\n",
                SyntaxKind::FunctionStatement,
            ),
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
    fn cst_function_with_modifiers() {
        // Test Public Function
        let source = "Public Function GetValue() As Integer\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("Public Function GetValue"));
    }

    #[test]
    fn cst_private_static_function() {
        // Test Private Static Function
        let source = "Private Static Function Calculate(x As Long) As Long\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("Private Static Function Calculate"));
    }

    #[test]
    fn cst_friend_function() {
        // Test Friend Function
        let source = "Friend Function ProcessData() As String\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("Friend Function ProcessData"));
    }

    #[test]
    fn function_with_line_continuation_in_params() {
        let source = r#"
Public Function Test( _
  ByVal x As Long _
) As String
    Test = "hello"
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();

        // Should have a FunctionStatement node
        assert!(
            debug.contains("FunctionStatement"),
            "Should be FunctionStatement"
        );
        // The function itself should not be parsed as a DimStatement
        // (although it may contain DimStatement nodes inside for variable declarations)
        assert!(
            debug.contains("  FunctionStatement@"),
            "Function should be at root level, not inside DimStatement"
        );
    }

    #[test]
    fn function_with_line_continuation_after_open_paren() {
        // This is the exact pattern from audiostation modArgs.bas argGetSwitchArg
        let source = r#"
Public Function argGetSwitchArg( _
  ByRef Switch As String, _
  Optional ByRef Position As Long, _
  Optional ByVal UseWildcard As Boolean _
) As String
Dim I&
argGetSwitchArg = ""
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();

        // Should have a FunctionStatement node
        assert!(
            debug.contains("FunctionStatement"),
            "Should be FunctionStatement"
        );
        // The function itself should not be parsed as a DimStatement
        assert!(
            debug.contains("  FunctionStatement@"),
            "Function should be at root level, not inside DimStatement"
        );
    }

    #[test]
    fn function_with_do_loop_before_end() {
        // Test that "End Function" after a DO loop is recognized correctly
        let source = r#"
Public Function Test(ByVal x As Long) As String
Dim i As Long
Do
    i = i + 1
Loop
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();

        assert!(
            debug.contains("FunctionStatement"),
            "Should have FunctionStatement"
        );
        assert!(
            !debug.contains("Unknown"),
            "Should not have any Unknown tokens"
        );
        assert!(
            debug.contains("  FunctionStatement@"),
            "Function should be at root level"
        );
    }

    #[test]
    fn function_with_line_continuation_in_if_condition() {
        // Test from audiostation modArgs.bas - line continuation in IF condition
        let source = r#"
Public Function argGetArgs(ByRef argv() As String, ByRef argc As Long, _
 Optional ByVal Args As String)
Dim strArgTemp As String
Do Until strArgTemp = ""
  If InStr(1, strArgTemp, Chr$(34)) <> 0 And _
     InStr(1, strArgTemp, Chr$(34)) < InStr(1, strArgTemp, " ") Then
    strArgTemp = ""
  End If
Loop
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();

        assert!(
            debug.contains("FunctionStatement"),
            "Should have FunctionStatement"
        );
        assert!(
            !debug.contains("Unknown"),
            "Should not have any Unknown tokens"
        );
        assert!(
            debug.contains("  FunctionStatement@"),
            "Function should be at root level"
        );
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
