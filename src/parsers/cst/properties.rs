use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    pub(super) fn parse_version_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::VersionStatement.to_raw());

        // VERSION
        self.consume_token();
        self.consume_whitespace();

        // Major.Minor
        // Consume until CLASS or Newline
        while !self.is_at_end()
            && !self.at_token(Token::ClassKeyword)
            && !self.at_token(Token::Newline)
            && !self.at_token(Token::ColonOperator)
        {
            self.consume_token();
        }

        // CLASS
        if self.at_token(Token::ClassKeyword) {
            self.consume_token();
        }

        self.consume_whitespace();
        self.consume_newline_or_colon();

        self.builder.finish_node();
    }

    pub(super) fn parse_properties_block(&mut self) {
        self.builder
            .start_node(SyntaxKind::PropertiesBlock.to_raw());

        // BEGIN
        self.consume_token();
        self.consume_whitespace();

        // Parse optional Type and Name (e.g. VB.Form Form1)
        // We check if we are at a newline or colon, if not, we assume there is a type/name
        if !self.at_token(Token::Newline) && !self.at_token(Token::ColonOperator) {
            // Type (e.g. VB.Form)
            self.builder.start_node(SyntaxKind::PropertiesType.to_raw());

            // Consume first part
            if self.is_identifier() || self.at_keyword() {
                self.consume_token_as_identifier();
            }

            // Consume dot parts
            while self.at_token(Token::PeriodOperator) {
                self.consume_token(); // .
                if self.is_identifier() || self.at_keyword() {
                    self.consume_token_as_identifier();
                }
            }
            self.builder.finish_node();

            self.consume_whitespace();

            // Name (e.g. Form1)
            if !self.at_token(Token::Newline) && !self.at_token(Token::ColonOperator) {
                self.builder.start_node(SyntaxKind::PropertiesName.to_raw());
                self.consume_token_as_identifier();
                self.builder.finish_node();
            }
        }

        self.consume_newline_or_colon();

        while !self.is_at_end() && !self.at_token(Token::EndKeyword) {
            self.consume_whitespace();

            if self.at_token(Token::EndKeyword) {
                break;
            }

            if self.at_token(Token::BeginKeyword) {
                // Check if this is BeginProperty or a nested control block
                // Peek ahead to see if next token is "Property"
                let next_pos = self.pos + 1;
                if next_pos < self.tokens.len() {
                    // Skip whitespace
                    let mut check_pos = next_pos;
                    while check_pos < self.tokens.len()
                        && self.tokens[check_pos].1 == Token::Whitespace
                    {
                        check_pos += 1;
                    }
                    // Check if it's "Property" keyword or identifier (tokenizer may vary)
                    let is_property = if check_pos < self.tokens.len() {
                        let text = self.tokens[check_pos].0;
                        self.tokens[check_pos].1 == Token::PropertyKeyword
                            || (self.tokens[check_pos].1 == Token::Identifier
                                && text.eq_ignore_ascii_case("Property"))
                    } else {
                        false
                    };

                    if is_property {
                        self.parse_property_group();
                    } else {
                        // Nested control block
                        self.parse_properties_block();
                    }
                } else {
                    self.parse_properties_block();
                }
            } else if self.is_identifier() {
                // Check if this is "BeginProperty" as a single identifier
                if self.pos < self.tokens.len()
                    && self.tokens[self.pos]
                        .0
                        .eq_ignore_ascii_case("BeginProperty")
                {
                    self.parse_property_group();
                } else {
                    self.parse_property();
                }
            } else if self.at_keyword() {
                self.parse_property();
            } else {
                // Skip unknown or newlines
                self.consume_token();
            }
        }

        // END
        if self.at_token(Token::EndKeyword) {
            self.consume_token();
        }

        self.consume_whitespace();
        self.consume_newline_or_colon();

        self.builder.finish_node();
    }

    fn parse_property(&mut self) {
        self.builder.start_node(SyntaxKind::Property.to_raw());

        // Key
        self.builder.start_node(SyntaxKind::PropertyKey.to_raw());
        self.consume_token(); // Identifier or Keyword
        self.builder.finish_node();

        self.consume_whitespace();

        // =
        if self.at_token(Token::EqualityOperator) {
            self.consume_token();
        }

        self.consume_whitespace();

        // Value
        self.builder.start_node(SyntaxKind::PropertyValue.to_raw());
        // Consume until newline
        while !self.is_at_end() && !self.at_token(Token::Newline) {
            self.consume_token();
        }
        self.builder.finish_node();

        self.consume_newline_or_colon();

        self.builder.finish_node();
    }

    fn parse_property_group(&mut self) {
        self.builder.start_node(SyntaxKind::PropertyGroup.to_raw());

        // BeginProperty - can be two tokens (Begin + Property) or one (BeginProperty)
        if self.at_token(Token::BeginKeyword) {
            self.consume_token(); // Begin
            self.consume_whitespace();
            self.consume_token(); // Property (keyword or identifier)
        } else if self.is_identifier()
            && self.pos < self.tokens.len()
            && self.tokens[self.pos]
                .0
                .eq_ignore_ascii_case("BeginProperty")
        {
            self.consume_token(); // BeginProperty
        } else {
            // Unexpected - just consume what's there
            self.consume_token();
        }
        self.consume_whitespace();

        // Property group name
        self.builder
            .start_node(SyntaxKind::PropertyGroupName.to_raw());
        if self.is_identifier() || self.at_keyword() {
            self.consume_token_as_identifier();
        }
        self.builder.finish_node();

        self.consume_whitespace();

        // Optional GUID and other tokens until newline - just consume them all
        while !self.is_at_end()
            && !self.at_token(Token::Newline)
            && !self.at_token(Token::ColonOperator)
        {
            self.consume_token();
        }

        self.consume_whitespace();
        self.consume_newline_or_colon();

        // Parse property group contents
        while !self.is_at_end() {
            self.consume_whitespace();

            // Check for EndProperty (can be two tokens or one)
            let is_end_property = if self.at_token(Token::EndKeyword) {
                let next_pos = self.pos + 1;
                if next_pos < self.tokens.len() {
                    let mut check_pos = next_pos;
                    while check_pos < self.tokens.len()
                        && self.tokens[check_pos].1 == Token::Whitespace
                    {
                        check_pos += 1;
                    }
                    check_pos < self.tokens.len()
                        && (self.tokens[check_pos].1 == Token::PropertyKeyword
                            || (self.tokens[check_pos].1 == Token::Identifier
                                && self.tokens[check_pos].0.eq_ignore_ascii_case("Property")))
                } else {
                    false
                }
            } else if self.is_identifier() && self.pos < self.tokens.len() {
                self.tokens[self.pos].0.eq_ignore_ascii_case("EndProperty")
            } else {
                false
            };

            if is_end_property {
                // Consume EndProperty
                if self.at_token(Token::EndKeyword) {
                    self.consume_token(); // End
                    self.consume_whitespace();
                    self.consume_token(); // Property
                } else {
                    self.consume_token(); // EndProperty
                }
                self.consume_whitespace();
                self.consume_newline_or_colon();
                break;
            }

            // Check for nested BeginProperty (can be two tokens or one)
            let is_begin_property = if self.at_token(Token::BeginKeyword) {
                let next_pos = self.pos + 1;
                if next_pos < self.tokens.len() {
                    let mut check_pos = next_pos;
                    while check_pos < self.tokens.len()
                        && self.tokens[check_pos].1 == Token::Whitespace
                    {
                        check_pos += 1;
                    }
                    check_pos < self.tokens.len()
                        && (self.tokens[check_pos].1 == Token::PropertyKeyword
                            || (self.tokens[check_pos].1 == Token::Identifier
                                && self.tokens[check_pos].0.eq_ignore_ascii_case("Property")))
                } else {
                    false
                }
            } else if self.is_identifier() && self.pos < self.tokens.len() {
                self.tokens[self.pos]
                    .0
                    .eq_ignore_ascii_case("BeginProperty")
            } else {
                false
            };

            if is_begin_property {
                self.parse_property_group();
                continue;
            }

            // Parse regular property
            if self.is_identifier() || self.at_keyword() {
                self.parse_property();
            } else if !self.at_token(Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node();
    }

    fn consume_newline_or_colon(&mut self) {
        if self.at_token(Token::Newline) || self.at_token(Token::ColonOperator) {
            self.consume_token();
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;

    #[test]
    fn class_parsing() {
        let source = r#"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Test()
    Dim x As Integer
    x = 1
End Sub
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Whitespace,
                ClassKeyword,
                Newline,
            },
            PropertiesBlock {
                BeginKeyword,
                Newline,
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("MultiUse"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        SubtractionOperator,
                        IntegerLiteral ("1"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("Persistable"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        IntegerLiteral ("0"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("DataBindingBehavior"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        IntegerLiteral ("0"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("DataSourceBehavior"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        IntegerLiteral ("0"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("MTSTransactionMode"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        IntegerLiteral ("0"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                EndKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Name"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"Class1\""),
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_GlobalNameSpace"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Creatable"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                TrueKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_PredeclaredId"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Exposed"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            OptionStatement {
                OptionKeyword,
                Whitespace,
                ExplicitKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("x"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn form_parsing() {
        let source = r#"VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "Hello"
End Sub
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Newline,
            },
            PropertiesBlock {
                BeginKeyword,
                Whitespace,
                PropertiesType {
                    Identifier ("VB"),
                    PeriodOperator,
                    Identifier ("Form"),
                },
                Whitespace,
                PropertiesName {
                    Identifier ("Form1"),
                },
                Whitespace,
                Newline,
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("Caption"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        StringLiteral ("\"Form1\""),
                    },
                    Newline,
                },
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("ClientHeight"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        IntegerLiteral ("3195"),
                    },
                    Newline,
                },
                Whitespace,
                PropertiesBlock {
                    BeginKeyword,
                    Whitespace,
                    PropertiesType {
                        Identifier ("VB"),
                        PeriodOperator,
                        Identifier ("CommandButton"),
                    },
                    Whitespace,
                    PropertiesName {
                        Identifier ("Command1"),
                    },
                    Whitespace,
                    Newline,
                    Whitespace,
                    Property {
                        PropertyKey {
                            Identifier ("Caption"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        PropertyValue {
                            StringLiteral ("\"Command1\""),
                        },
                        Newline,
                    },
                    Whitespace,
                    EndKeyword,
                    Newline,
                },
                EndKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Name"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"Form1\""),
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_GlobalNameSpace"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Creatable"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_PredeclaredId"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                TrueKeyword,
                Newline,
            },
            AttributeStatement {
                AttributeKeyword,
                Whitespace,
                Identifier ("VB_Exposed"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                FalseKeyword,
                Newline,
            },
            OptionStatement {
                OptionKeyword,
                Whitespace,
                ExplicitKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Command1_Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Hello\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
