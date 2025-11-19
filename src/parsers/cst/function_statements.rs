//! Function statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Function statements with syntax:
//!
//! \[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]
//! \[ statements \]
//! \[ name = expression \]
//! \[ Exit Function \]
//! \[ statements \]
//! \[ name = expression \]
//! End Function
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn function_distinguishes_declarations_from_functions() {
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
    fn function_all_modifier_combinations() {
        // Test all valid function modifier combinations
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
    fn function_with_modifiers() {
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
    fn function_private_static() {
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
    fn function_friend() {
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
    fn function_simple_no_params() {
        // Test simple function with no parameters
        let source = "Function GetValue() As Integer\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
    }

    #[test]
    fn function_with_return_value() {
        // Test function with return value assignment
        let source = "Function GetValue() As Integer\n    GetValue = 42\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("GetValue = 42"));
    }

    #[test]
    fn function_with_exit_function() {
        // Test function with Exit Function statement
        let source = "Function IsValid(x As Integer) As Boolean\n    If x < 0 Then\n        Exit Function\n    End If\n    IsValid = True\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("Exit Function"));
    }

    #[test]
    fn function_no_return_type() {
        // Test function without explicit return type (defaults to Variant)
        let source = "Function GetData()\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
    }

    #[test]
    fn function_with_multiple_params() {
        // Test function with multiple parameters
        let source = "Function Add(ByVal x As Long, ByVal y As Long) As Long\nEnd Function\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::FunctionStatement);
        }
        assert!(cst.text().contains("ByVal x As Long, ByVal y As Long"));
    }
}
