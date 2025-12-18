//! Parameter list parsing for VB6 CST.
//!
//! This module handles parsing of parameter lists in VB6 procedures:
//! - Function parameter lists
//! - Sub parameter lists
//! - Property parameter lists

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a parameter list: (param1 As Type, param2 As Type)
    ///
    /// VB6 parameter list syntax supports:
    /// - [Optional] [`ByVal` | `ByRef`] [`ParamArray`] varname[()] [As type] [= defaultvalue]
    ///
    /// This parser handles nested parentheses for array parameters and default values.
    pub(super) fn parse_parameter_list(&mut self) {
        self.builder.start_node(SyntaxKind::ParameterList.to_raw());

        // Consume "("
        self.consume_token();

        loop {
            self.consume_whitespace();

            if self.at_token(Token::RightParenthesis) || self.is_at_end() {
                break;
            }

            // Optional
            if self.at_token(Token::OptionalKeyword) {
                self.consume_token();
                self.consume_whitespace();
            }

            // ByVal / ByRef
            if self.at_token(Token::ByValKeyword) || self.at_token(Token::ByRefKeyword) {
                self.consume_token();
                self.consume_whitespace();
            }

            // ParamArray
            if self.at_token(Token::ParamArrayKeyword) {
                self.consume_token();
                self.consume_whitespace();
            }

            // Variable name
            if self.at_token(Token::Identifier) {
                self.consume_token();
            } else {
                // Error recovery
                break;
            }

            self.consume_whitespace();

            // Array parens ()
            if self.at_token(Token::LeftParenthesis) {
                self.consume_token();
                self.consume_whitespace();
                if self.at_token(Token::RightParenthesis) {
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // As Type
            if self.at_token(Token::AsKeyword) {
                self.consume_token();
                self.consume_whitespace();
                // Type name
                self.consume_token();
                while self.at_token(Token::PeriodOperator) {
                    self.consume_token();
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // Default value
            if self.at_token(Token::EqualityOperator) {
                self.consume_token();
                self.consume_whitespace();
                self.parse_expression();
            }

            self.consume_whitespace();

            if self.at_token(Token::Comma) {
                self.consume_token();
            } else {
                break;
            }
        }

        // Consume ")"
        if self.at_token(Token::RightParenthesis) {
            self.consume_token();
        }

        self.builder.finish_node(); // ParameterList
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn parameter_list_empty() {
        let source = r"
Sub Test()
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SubStatement"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_single_parameter() {
        let source = r"
Sub Test(x As Integer)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SubStatement"));
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("Integer"));
    }

    #[test]
    fn parameter_list_multiple_parameters() {
        let source = r"
Function Calculate(x As Integer, y As Integer, z As Integer) As Integer
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FunctionStatement"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_with_byval() {
        let source = r"
Sub Process(ByVal value As Long)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("ByValKeyword"));
    }

    #[test]
    fn parameter_list_with_byref() {
        let source = r"
Sub Modify(ByRef value As Long)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("ByRefKeyword"));
    }

    #[test]
    fn parameter_list_with_optional() {
        let source = r"
Sub Test(Optional x As Integer)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("OptionalKeyword"));
    }

    #[test]
    fn parameter_list_with_default_value() {
        let source = r"
Sub Test(Optional x As Integer = 10)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("OptionalKeyword"));
    }

    #[test]
    fn parameter_list_with_array() {
        let source = r"
Sub ProcessArray(arr() As Integer)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("arr"));
    }

    #[test]
    fn parameter_list_with_paramarray() {
        let source = r"
Sub VarArgs(ParamArray args() As Variant)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("ParamArrayKeyword"));
    }

    #[test]
    fn parameter_list_mixed_modifiers() {
        let source = r#"
Sub Test(ByVal x As Integer, ByRef y As Long, Optional z As String = "")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("ByValKeyword"));
        assert!(debug.contains("ByRefKeyword"));
        assert!(debug.contains("OptionalKeyword"));
    }

    #[test]
    fn parameter_list_with_object_type() {
        let source = r"
Sub SetObject(obj As Object)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("ObjectKeyword"));
    }

    #[test]
    fn parameter_list_with_variant() {
        let source = r"
Function ProcessData(data As Variant) As Boolean
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("VariantKeyword"));
    }

    #[test]
    fn parameter_list_with_line_continuation() {
        let source = r"
Public Function Test( _
  ByVal x As Long _
) As String
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FunctionStatement"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_multiple_with_line_continuation() {
        let source = r"
Public Function Process( _
  ByRef Switch As String, _
  Optional ByRef Position As Long, _
  Optional ByVal UseWildcard As Boolean _
) As String
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FunctionStatement"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_preserves_whitespace() {
        let source = r"
Sub Test(  x   As   Integer  ,  y   As   Long  )
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn parameter_list_nested_parentheses() {
        let source = r"
Sub Test(arr() As Integer, Optional index As Long = (5 + 3))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_custom_type() {
        let source = r"
Sub Process(emp As Employee)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
        assert!(debug.contains("Employee"));
    }

    #[test]
    fn parameter_list_no_type_specified() {
        let source = r"
Sub Test(x, y, z)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_in_property_get() {
        let source = r"
Property Get Item(index As Long) As Variant
End Property
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn parameter_list_in_property_let() {
        let source = r"
Property Let Value(newValue As Long)
End Property
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("ParameterList"));
    }
}
