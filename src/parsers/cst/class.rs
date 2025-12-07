use crate::language::VB6Token;
use crate::parsers::SyntaxKind;
use crate::parsers::cst::Parser;

impl Parser<'_> {
    pub(super) fn parse_version_statement(&mut self) {
        self.builder.start_node(SyntaxKind::VersionStatement.to_raw());
        
        // VERSION
        self.consume_token(); 
        self.consume_whitespace();
        
        // Major.Minor
        // This might be tokenized as FloatLiteral or Integer Period Integer
        // Let's just consume until CLASS
        
        while !self.is_at_end() && !self.at_token(VB6Token::ClassKeyword) {
            self.consume_token();
        }
        
        // CLASS
        if self.at_token(VB6Token::ClassKeyword) {
            self.consume_token();
        }
        
        self.consume_whitespace();
        self.consume_newline_or_colon();
        
        self.builder.finish_node();
    }

    pub(super) fn parse_properties_block(&mut self) {
        self.builder.start_node(SyntaxKind::PropertiesBlock.to_raw());
        
        // BEGIN
        self.consume_token();
        self.consume_whitespace();
        self.consume_newline_or_colon();
        
        while !self.is_at_end() && !self.at_token(VB6Token::EndKeyword) {
            self.consume_whitespace();
            
            if self.at_token(VB6Token::EndKeyword) {
                break;
            }
            
            if self.at_token(VB6Token::BeginKeyword) {
                // Nested block
                self.parse_properties_block();
            } else if self.is_identifier() || self.at_keyword() {
                self.parse_property();
            } else {
                // Skip unknown or newlines
                self.consume_token();
            }
        }
        
        // END
        if self.at_token(VB6Token::EndKeyword) {
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
        if self.at_token(VB6Token::EqualityOperator) {
            self.consume_token();
        }
        
        self.consume_whitespace();
        
        // Value
        self.builder.start_node(SyntaxKind::PropertyValue.to_raw());
        // Consume until newline
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        self.builder.finish_node();
        
        self.consume_newline_or_colon();
        
        self.builder.finish_node();
    }

    fn consume_newline_or_colon(&mut self) {
        if self.at_token(VB6Token::Newline) || self.at_token(VB6Token::ColonOperator) {
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
        assert!(cst.contains_kind(SyntaxKind::VersionStatement), "Should contain VersionStatement");
        
        // Check for PropertiesBlock
        assert!(cst.contains_kind(SyntaxKind::PropertiesBlock), "Should contain PropertiesBlock");
        
        // Check for AttributeStatement
        assert!(cst.contains_kind(SyntaxKind::AttributeStatement), "Should contain AttributeStatement");
        
        // Check for SubStatement
        assert!(cst.contains_kind(SyntaxKind::SubStatement), "Should contain SubStatement");
        
        // Check text preservation
        assert_eq!(cst.text(), input);
    }
}
