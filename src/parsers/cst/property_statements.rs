//! Property statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Property statements:
//! - Property Get
//! - Property Let
//! - Property Set

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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

        // Consume "Property" keyword
        self.consume_token();

        // Consume any whitespace after "Property"
        self.consume_whitespace();

        // Consume Get/Let/Set keyword
        if self.at_token(VB6Token::GetKeyword)
            || self.at_token(VB6Token::LetKeyword)
            || self.at_token(VB6Token::SetKeyword)
        {
            self.consume_token();
        }

        // Consume any whitespace after Get/Let/Set
        self.consume_whitespace();

        // Consume property name (keywords can be used as property names in VB6)
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
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

        // Parse body until "End Property"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::PropertyKeyword)
        });

        // Consume "End Property" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Property"
            self.consume_whitespace();

            // Consume "Property"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // PropertyStatement
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::ConcreteSyntaxTree;

    #[test]
    fn property_get_simple() {
        let source = r#"
Property Get Name() As String
    Name = m_name
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PropertyKeyword"));
        assert!(debug.contains("GetKeyword"));
    }

    #[test]
    fn property_let_simple() {
        let source = r#"
Property Let Name(ByVal newName As String)
    m_name = newName
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PropertyKeyword"));
        assert!(debug.contains("LetKeyword"));
    }

    #[test]
    fn property_set_simple() {
        let source = r#"
Property Set Container(glistNN As gList)
    Set glistN = glistNN
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PropertyKeyword"));
        assert!(debug.contains("SetKeyword"));
    }

    #[test]
    fn property_set_with_object() {
        let source = r#"
Property Set Callback(ByRef newObj As InterPress)
    Set mCallback = newObj
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("SetKeyword"));
        assert!(debug.contains("SetStatement")); // Set statement inside the property
    }

    #[test]
    fn property_get_public() {
        let source = r#"
Public Property Get Value() As Long
    Value = m_value
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PublicKeyword"));
        assert!(debug.contains("GetKeyword"));
    }

    #[test]
    fn property_let_private() {
        let source = r#"
Private Property Let Count(ByVal newCount As Integer)
    m_count = newCount
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("LetKeyword"));
    }

    #[test]
    fn property_set_friend() {
        let source = r#"
Friend Property Set objref(RHS As Object)
    Set m_objref = RHS
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("FriendKeyword"));
        assert!(debug.contains("SetKeyword"));
    }

    #[test]
    fn property_get_with_parameters() {
        let source = r#"
Public Property Get Item(index As Long) As Variant
    Item = m_items(index)
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("GetKeyword"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn property_set_with_index_parameter() {
        let source = r#"
Public Property Set item(curitem As Long, item As Variant)
    Set m_items(curitem) = item
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("SetKeyword"));
        assert!(debug.contains("ParameterList"));
    }

    #[test]
    fn property_get_with_if_statement() {
        let source = r#"
Public Property Get CustomColor(i As Integer) As Long
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("GetKeyword"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn property_let_with_if_statement() {
        let source = r#"
Public Property Let CustomColor(i As Integer, iValue As Long)
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("LetKeyword"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn property_get_no_parameters() {
        let source = r#"
Property Get APIReturn() As Long
    APIReturn = m_lApiReturn
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("GetKeyword"));
    }

    #[test]
    fn property_set_preserves_whitespace() {
        let source = r#"
Property   Set   Container  (  glistNN   As   gList  )
    Set glistN = glistNN
End   Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_properties_in_class() {
        let source = r#"
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
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let property_count = debug.matches("PropertyStatement").count();
        assert_eq!(property_count, 4);
    }

    #[test]
    fn property_get_returns_object() {
        let source = r#"
Property Get Callback() As InterPress
    Set Callback = mCallback
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("GetKeyword"));
        assert!(debug.contains("SetStatement")); // Set used for object return
    }

    #[test]
    fn property_with_exit_property() {
        let source = r#"
Property Get Test() As String
    If m_value = "" Then Exit Property
    Test = m_value
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("ExitStatement"));
        assert!(debug.contains("ExitKeyword"));
    }

    #[test]
    fn property_static() {
        let source = r#"
Public Static Property Get Counter() As Long
    Static count As Long
    count = count + 1
    Counter = count
End Property
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("PublicKeyword"));
        assert!(debug.contains("StaticKeyword"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PropertyStatement"));
        assert!(debug.contains("LetKeyword"));
        assert!(debug.contains("IfStatement"));
    }
}
