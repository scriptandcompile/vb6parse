//! Integration test for CST parsing functionality

use vb6parse::language::VB6Token;
use vb6parse::parsers::{parse, ConcreteSyntaxTree, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn cst_basic_parsing() {
    let code = "Sub Test()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Now has 1 child: the SubStatement node (with EOF inside as Unknown)
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Sub Test()\n");
}

#[test]
fn cst_complex_code() {
    let code = r#"Private Sub Calculate(ByVal x As Integer)
    Dim result As Long
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let text = cst.text();
    assert!(text.contains("Private Sub Calculate"));
    assert!(text.contains("ByVal x As Integer"));
    assert!(text.contains("Dim result As Long"));
    assert!(text.contains("End Sub"));
}

#[test]
fn cst_preserves_all_whitespace() {
    let code = "Sub  Test (  )\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
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
    assert_eq!(SyntaxKind::from(VB6Token::IfKeyword), SyntaxKind::IfKeyword);
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
    assert_eq!(SyntaxKind::from(VB6Token::Number), SyntaxKind::Number);
}

#[test]
fn public_api_does_not_expose_rowan() {
    // This test verifies that we can use the CST API without ever
    // importing or referring to rowan types

    let code = "Sub Test()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");

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

#[test]
fn parse_empty_stream() {
    let code = "";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
}

#[test]
fn parse_simple_tokens() {
    let code = "Sub Main()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 1 child: the SubStatement node
    assert_eq!(cst.child_count(), 1);
    assert!(cst.text().contains("Sub Main()"));
}

#[test]
fn parse_attribute_statement() {
    let code = "Attribute VB_Name = \"modTest\"\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1); // AttributeStatement
    assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
    let attr_statements = cst.find_children_by_kind(SyntaxKind::AttributeStatement);
    assert_eq!(attr_statements.len(), 1);
    assert_eq!(attr_statements[0].text, "Attribute VB_Name = \"modTest\"\n");
}
