//! Property statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Property statements:
//! - Property Get
//! - Property Let
//! - Property Set

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Property statement (Property Get, Property Let, or Property Set).
    ///
    /// VB6 Property statement syntax:
    /// - [Public | Private | Friend] [Static] Property Get name [(arglist)] [As type]
    /// - [Public | Private | Friend] [Static] Property Let name ([arglist,] value)
    /// - [Public | Private | Friend] [Static] Property Set name ([arglist,] value)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-get-statement)
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-let-statement)
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-set-statement)
    pub(super) fn parse_property_statement(&mut self) {
        // if we are now parsing a property statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::PropertyStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

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

        // Consume "Property" keyword
        self.consume_token();

        // Consume any whitespace after "Property"
        self.consume_whitespace();

        // Consume Get/Let/Set keyword
        if self.at_token(Token::GetKeyword)
            || self.at_token(Token::LetKeyword)
            || self.at_token(Token::SetKeyword)
        {
            self.consume_token();
        }

        // Consume any whitespace after Get/Let/Set
        self.consume_whitespace();

        // Consume property name (keywords can be used as property names in VB6)
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

        // Consume everything until newline (includes "As Type" if present)
        self.consume_until_after(Token::Newline);

        // Parse body until "End Property"
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::PropertyKeyword)
        });

        // Consume "End Property" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Property"
            self.consume_whitespace();

            // Consume "Property"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // PropertyStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::parsers::ConcreteSyntaxTree;
    #[test]
    fn property_get_simple() {
        let source = r"
Property Get Name() As String
    Name = m_name
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
                Identifier ("Name"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("m_name"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_let_simple() {
        let source = r"
Property Let Name(ByVal newName As String)
    m_name = newName
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
                Identifier ("Name"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newName"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_name"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("newName"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_set_simple() {
        let source = r"
Property Set Container(glistNN As gList)
    Set glistN = glistNN
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                SetKeyword,
                Whitespace,
                Identifier ("Container"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("glistNN"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("gList"),
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("glistN"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("glistNN"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_set_with_object() {
        let source = r"
Property Set Callback(ByRef newObj As InterPress)
    Set mCallback = newObj
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                SetKeyword,
                Whitespace,
                Identifier ("Callback"),
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("newObj"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("InterPress"),
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("mCallback"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newObj"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_get_public() {
        let source = r"
Public Property Get Value() As Long
    Value = m_value
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("m_value"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_let_private() {
        let source = r"
Private Property Let Count(ByVal newCount As Integer)
    m_count = newCount
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PrivateKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Count"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newCount"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_count"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("newCount"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_set_friend() {
        let source = r"
Friend Property Set objref(RHS As Object)
    Set m_objref = RHS
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                FriendKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                SetKeyword,
                Whitespace,
                Identifier ("objref"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("RHS"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ObjectKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("m_objref"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("RHS"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_get_with_parameters() {
        let source = r"
Public Property Get Item(index As Long) As Variant
    Item = m_items(index)
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
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
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Item"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("m_items"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("index"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_set_with_index_parameter() {
        let source = r"
Public Property Set item(curitem As Long, item As Variant)
    Set m_items(curitem) = item
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                SetKeyword,
                Whitespace,
                Identifier ("item"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("curitem"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("item"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    VariantKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("m_items"),
                        LeftParenthesis,
                        Identifier ("curitem"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("item"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_get_with_if_statement() {
        let source = r"
Public Property Get CustomColor(i As Integer) As Long
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("CustomColor"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("i"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("fNotFirst"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            BooleanLiteralExpression {
                                FalseKeyword,
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        Identifier ("InitColors"),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("15"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("CustomColor"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("alCustom"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("i"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("CustomColor"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_let_with_if_statement() {
        let source = r"
Public Property Let CustomColor(i As Integer, iValue As Long)
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("CustomColor"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("i"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("iValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("fNotFirst"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            BooleanLiteralExpression {
                                FalseKeyword,
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        Identifier ("InitColors"),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("15"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                CallExpression {
                                    Identifier ("alCustom"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("i"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("iValue"),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_get_no_parameters() {
        let source = r"
Property Get APIReturn() As Long
    APIReturn = m_lApiReturn
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
                Identifier ("APIReturn"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("APIReturn"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("m_lApiReturn"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_set_preserves_whitespace() {
        let source = r"
Property   Set   Container  (  glistNN   As   gList  )
    Set glistN = glistNN
End   Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                SetKeyword,
                Whitespace,
                Identifier ("Container"),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Identifier ("glistNN"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("gList"),
                    Whitespace,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("glistN"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("glistNN"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn multiple_properties_in_class() {
        let source = r"
Private m_name As String
Private m_value As Long

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal newName As String)
    m_name = newName
End Property

Public Property Get Value() As Long
    Value = m_value
End Property

Public Property Let Value(ByVal newValue As Long)
    m_value = newValue
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_name"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_value"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Name"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("m_name"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Name"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newName"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_name"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("newName"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("m_value"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("newValue"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_get_returns_object() {
        let source = r"
Property Get Callback() As InterPress
    Set Callback = mCallback
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
                Identifier ("Callback"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("InterPress"),
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("Callback"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("mCallback"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_with_exit_property() {
        let source = r#"
Property Get Test() As String
    If m_value = "" Then Exit Property
    Test = m_value
End Property
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("m_value"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        ExitStatement {
                            ExitKeyword,
                            Whitespace,
                            PropertyKeyword,
                            Newline,
                        },
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("Test"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("m_value"),
                            },
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        PropertyKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn property_static() {
        let source = r"
Public Static Property Get Counter() As Long
    Static count As Long
    count = count + 1
    Counter = count
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Counter"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        StaticKeyword,
                        Whitespace,
                        Identifier ("count"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("count"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("count"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Counter"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("count"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn property_let_with_multiple_statements() {
        let source = r#"
Public Property Let Caption(myCap As String)
    mCaptext = myCap
    If Not glistN Is Nothing Then
        If glistN.CenterText Then
            glistN.list(0) = mCaptext
        Else
            glistN.list(0) = "  " + mCaptext
        End If
        glistN.ShowMe
    End If
End Property
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Caption"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("myCap"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("mCaptext"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("myCap"),
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("glistN"),
                                },
                                Whitespace,
                                IsKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("Nothing"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                MemberAccessExpression {
                                    Identifier ("glistN"),
                                    PeriodOperator,
                                    Identifier ("CenterText"),
                                },
                                Whitespace,
                                ThenKeyword,
                                Newline,
                                StatementList {
                                    Whitespace,
                                    AssignmentStatement {
                                        CallExpression {
                                            MemberAccessExpression {
                                                Identifier ("glistN"),
                                                PeriodOperator,
                                                Identifier ("list"),
                                            },
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("0"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("mCaptext"),
                                        },
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                ElseClause {
                                    ElseKeyword,
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        AssignmentStatement {
                                            CallExpression {
                                                MemberAccessExpression {
                                                    Identifier ("glistN"),
                                                    PeriodOperator,
                                                    Identifier ("list"),
                                                },
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("0"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            BinaryExpression {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"  \""),
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("mCaptext"),
                                                },
                                            },
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                },
                                EndKeyword,
                                Whitespace,
                                IfKeyword,
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("glistN"),
                                PeriodOperator,
                                Identifier ("ShowMe"),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }
}
