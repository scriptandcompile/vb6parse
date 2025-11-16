//! Integration test for CST parsing functionality

use vb6parse::language::VB6Token;
use vb6parse::parsers::{parse, SyntaxKind, ConcreteSyntaxTree};
use vb6parse::tokenstream::TokenStream;

#[test]
fn cst_basic_parsing() {
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Test", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Now has 1 child: the SubStatement node
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Sub Test()\n");
}

#[test]
fn cst_with_comments() {
    let tokens = vec![
        ("' This is a comment\n", VB6Token::EndOfLineComment),
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Main", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    // Now has 2 children: the comment (consumed at top level) and the SubStatement
    assert_eq!(cst.child_count(), 2);
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));
}

#[test]
fn cst_complex_code() {
    let tokens = vec![
        ("Private", VB6Token::PrivateKeyword),
        (" ", VB6Token::Whitespace),
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Calculate", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        ("ByVal", VB6Token::ByValKeyword),
        (" ", VB6Token::Whitespace),
        ("x", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),
        ("As", VB6Token::AsKeyword),
        (" ", VB6Token::Whitespace),
        ("Integer", VB6Token::IntegerKeyword),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
        ("    ", VB6Token::Whitespace),
        ("Dim", VB6Token::DimKeyword),
        (" ", VB6Token::Whitespace),
        ("result", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),
        ("As", VB6Token::AsKeyword),
        (" ", VB6Token::Whitespace),
        ("Long", VB6Token::LongKeyword),
        ("\n", VB6Token::Newline),
        ("End", VB6Token::EndKeyword),
        (" ", VB6Token::Whitespace),
        ("Sub", VB6Token::SubKeyword),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    let text = cst.text();
    assert!(text.contains("Private Sub Calculate"));
    assert!(text.contains("ByVal x As Integer"));
    assert!(text.contains("Dim result As Long"));
    assert!(text.contains("End Sub"));
}

#[test]
fn cst_preserves_all_whitespace() {
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        ("  ", VB6Token::Whitespace), // Two spaces
        ("Test", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),  // One space
        ("(", VB6Token::LeftParentheses),
        ("  ", VB6Token::Whitespace), // Two spaces
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    // Verify that all whitespace is preserved exactly
    assert_eq!(cst.text(), "Sub  Test (  )\n");
}

#[test]
fn syntax_kind_conversions() {
    // Test keyword conversions
    assert_eq!(
        SyntaxKind::from(VB6Token::FunctionKeyword),
        SyntaxKind::FunctionKeyword
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::IfKeyword),
        SyntaxKind::IfKeyword
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::ForKeyword),
        SyntaxKind::ForKeyword
    );

    // Test operators
    assert_eq!(
        SyntaxKind::from(VB6Token::AdditionOperator),
        SyntaxKind::AdditionOperator
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::EqualityOperator),
        SyntaxKind::EqualityOperator
    );

    // Test literals
    assert_eq!(
        SyntaxKind::from(VB6Token::StringLiteral),
        SyntaxKind::StringLiteral
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::Number),
        SyntaxKind::Number
    );
}

#[test]
fn empty_token_stream() {
    let tokens = vec![];
    let token_stream = TokenStream::new("empty.bas".to_string(), tokens);
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
    assert_eq!(cst.text(), "");
}


#[test]
fn public_api_does_not_expose_rowan() {
    // This test verifies that we can use the CST API without ever
    // importing or referring to rowan types
    
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Test", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];
    
    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    
    // Parse using the public API
    let cst: ConcreteSyntaxTree = parse(token_stream);
    
    // Use all public methods
    let _kind: SyntaxKind = cst.root_kind();
    let _text: String = cst.text();
    let _count: usize = cst.child_count();
    let _debug: String = cst.debug_tree();
    
    // Verify the CST can be cloned (Clone is derived)
    let _cst_clone = cst.clone();
    
    // Verify equality checks work (PartialEq, Eq are derived)
    assert_eq!(cst, _cst_clone);
    
    // If this test compiles and runs without importing rowan,
    // then we've successfully hidden all rowan types!
}

#[test]
fn syntax_kind_public_interface() {
    // Verify SyntaxKind can be used without rowan
    
    let kind = SyntaxKind::from(VB6Token::SubKeyword);
    assert_eq!(kind, SyntaxKind::SubKeyword);
    
    // Verify Debug trait works
    let debug_str = format!("{:?}", kind);
    assert!(debug_str.contains("SubKeyword"));
    
    // Verify comparisons work
    assert_eq!(kind, SyntaxKind::SubKeyword);
    assert_ne!(kind, SyntaxKind::FunctionKeyword);
    
    // Verify Clone works
    let _cloned = kind.clone();
}
