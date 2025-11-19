//! Integration test for CST parsing functionality

use vb6parse::language::VB6Token;
use vb6parse::parsers::{ConcreteSyntaxTree, SyntaxKind};

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
fn parse_empty_stream() {
    let source = "";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
}

#[test]
fn parse_attribute_statement() {
    let source = "Attribute VB_Name = \"modTest\"\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1); // AttributeStatement
    assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
    let attr_statements = cst.find_children_by_kind(SyntaxKind::AttributeStatement);
    assert_eq!(attr_statements.len(), 1);
    assert_eq!(attr_statements[0].text, "Attribute VB_Name = \"modTest\"\n");
}
