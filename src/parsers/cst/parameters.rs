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
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn parameter_list_empty() {
        let source = r"
Sub Test()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_single_parameter() {
        let source = r"
Sub Test(x As Integer)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_multiple_parameters() {
        let source = r"
Function Calculate(x As Integer, y As Integer, z As Integer) As Integer
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Calculate"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("z"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_byval() {
        let source = r"
Sub Process(ByVal value As Long)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Process"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_byref() {
        let source = r"
Sub Modify(ByRef value As Long)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Modify"),
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_optional() {
        let source = r"
Sub Test(Optional x As Integer)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    OptionalKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_default_value() {
        let source = r"
Sub Test(Optional x As Integer = 10)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    OptionalKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("10"),
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_array() {
        let source = r"
Sub ProcessArray(arr() As Integer)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessArray"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("arr"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_paramarray() {
        let source = r"
Sub VarArgs(ParamArray args() As Variant)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("VarArgs"),
                ParameterList {
                    LeftParenthesis,
                    ParamArrayKeyword,
                    Whitespace,
                    Identifier ("args"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    VariantKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_mixed_modifiers() {
        let source = r#"
Sub Test(ByVal x As Integer, ByRef y As Long, Optional z As String = "")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    Identifier ("z"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_object_type() {
        let source = r"
Sub SetObject(obj As Object)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("SetObject"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("obj"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ObjectKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_variant() {
        let source = r"
Function ProcessData(data As Variant) As Boolean
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    VariantKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_with_line_continuation() {
        let source = r"
Public Function Test( _
  ByVal x As Long _
) As String
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Whitespace,
                    Underscore,
                    Newline,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Process"),
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Switch"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Position"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("UseWildcard"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    BooleanKeyword,
                    Whitespace,
                    Underscore,
                    Newline,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_preserves_whitespace() {
        let source = r"
Sub Test(  x   As   Integer  ,  y   As   Long  )
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Whitespace,
                    Comma,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Whitespace,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_nested_parentheses() {
        let source = r"
Sub Test(arr() As Integer, Optional index As Long = (5 + 3))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("arr"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    Identifier ("index"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            NumericLiteralExpression {
                                IntegerLiteral ("5"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("3"),
                            },
                        },
                        RightParenthesis,
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_custom_type() {
        let source = r"
Sub Process(emp As Employee)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Process"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("emp"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Employee"),
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_no_type_specified() {
        let source = r"
Sub Test(x, y, z)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Comma,
                    Whitespace,
                    Identifier ("y"),
                    Comma,
                    Whitespace,
                    Identifier ("z"),
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_in_property_get() {
        let source = r"
Property Get Item(index As Long) As Variant
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Item"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("index"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn parameter_list_in_property_let() {
        let source = r"
Property Let Value(newValue As Long)
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("newValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }
}
