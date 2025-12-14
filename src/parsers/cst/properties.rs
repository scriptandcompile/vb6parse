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
            } else {
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
    use crate::parsers::cst::parse;
    use crate::parsers::SyntaxKind;
    use crate::tokenize::tokenize;
    use crate::SourceStream;

    #[test]
    fn class_parsing() {
        let input = r#"VERSION 1.0 CLASS
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

        let mut source_stream = SourceStream::new("Class1.cls", input);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization failed");

        let cst = parse(token_stream);

        // Check structure
        assert_eq!(cst.root_kind(), SyntaxKind::Root);

        // Check for VersionStatement
        assert!(
            cst.contains_kind(SyntaxKind::VersionStatement),
            "Should contain VersionStatement"
        );

        // Check for PropertiesBlock
        assert!(
            cst.contains_kind(SyntaxKind::PropertiesBlock),
            "Should contain PropertiesBlock"
        );

        // Check for AttributeStatement
        assert!(
            cst.contains_kind(SyntaxKind::AttributeStatement),
            "Should contain AttributeStatement"
        );

        // Check for SubStatement
        assert!(
            cst.contains_kind(SyntaxKind::SubStatement),
            "Should contain SubStatement"
        );

        // Check text preservation
        assert_eq!(cst.text(), input);
    }

    #[test]
    fn form_parsing() {
        let input = r#"VERSION 5.00
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

        let mut source_stream = SourceStream::new("Form1.frm", input);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization failed");

        let cst = parse(token_stream);

        // Check structure
        assert_eq!(cst.root_kind(), SyntaxKind::Root);

        // Check for VersionStatement
        assert!(
            cst.contains_kind(SyntaxKind::VersionStatement),
            "Should contain VersionStatement"
        );

        // Check for PropertiesBlock
        assert!(
            cst.contains_kind(SyntaxKind::PropertiesBlock),
            "Should contain PropertiesBlock"
        );

        // Check for AttributeStatement
        assert!(
            cst.contains_kind(SyntaxKind::AttributeStatement),
            "Should contain AttributeStatement"
        );

        // Check for SubStatement
        assert!(
            cst.contains_kind(SyntaxKind::SubStatement),
            "Should contain SubStatement"
        );

        // Check text preservation
        assert_eq!(cst.text(), input);
    }
}
