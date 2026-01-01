//! Tests for keywords used as identifiers in various positions
//!
//! VB6 allows keywords to be used as identifiers (variable names, procedure names, etc.)
//! in most contexts. This test file verifies that keywords are properly converted to
//! Identifier tokens when they appear in identifier positions.

use vb6parse::parsers::{ConcreteSyntaxTree, SyntaxKind};

#[test]
fn keyword_as_sub_name() {
    let source = "Sub Text()\nEnd Sub\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    // "Text" should be converted to Identifier
    assert!(debug.contains("Identifier") && debug.contains("Text"));
}

#[test]
fn keyword_as_function_name() {
    let source = "Function Database() As String\nEnd Function\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    assert!(debug.contains("Identifier") && debug.contains("Database"));
}

#[test]
fn keyword_as_property_name() {
    let source = "Property Get Binary() As Integer\nEnd Property\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    assert!(debug.contains("Identifier") && debug.contains("Binary"));
}

#[test]
fn keyword_as_variable_in_assignment() {
    let source = "text = \"hello\"\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let children = cst.children();
    assert_eq!(children[0].kind(), SyntaxKind::AssignmentStatement);
    // text is now wrapped in IdentifierExpression
    assert_eq!(
        children[0].children()[0].kind(),
        SyntaxKind::IdentifierExpression
    );
    assert!(children[0].children()[0].text().contains("text"));
}

#[test]
fn keyword_as_property_in_assignment() {
    let source = "obj.text = \"hello\"\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let children = cst.children();
    assert_eq!(children[0].kind(), SyntaxKind::AssignmentStatement);
    // obj.text is now wrapped in MemberAccessExpression
    assert_eq!(
        children[0].children()[0].kind(),
        SyntaxKind::MemberAccessExpression
    );
    // The member access should contain "obj" and "text"
    assert!(children[0].children()[0].text().contains("obj"));
    assert!(children[0].children()[0].text().contains("text"));
}

#[test]
fn multiple_keywords_as_identifiers() {
    let source = r#"
database = "mydb.mdb"
text = "hello"
obj.binary = True
"#;
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    // All should be converted to Identifiers
    assert!(debug.contains("AssignmentStatement"));
    // Count how many times Identifier appears - should be at least 3
    // (database, text, obj - binary may not be counted separately in new structure)
    let identifier_count = debug.matches("Identifier").count();
    assert!(
        identifier_count >= 3,
        "Expected at least 3 identifiers, found {identifier_count}"
    );
}

#[test]
fn keyword_as_enum_name() {
    let source = "Enum Random\n    Value1\n    Value2\nEnd Enum\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    assert!(debug.contains("EnumStatement"));
    // "Random" should be converted to Identifier
    assert!(debug.contains("Identifier") && debug.contains("Random"));
}

#[test]
fn keyword_after_keyword_converted() {
    // Even when a keyword follows another keyword in procedure definition
    let source = "Sub Output()\nEnd Sub\n";
    let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

    let debug = cst.debug_tree();
    assert!(debug.contains("SubStatement"));
    assert!(debug.contains("Identifier") && debug.contains("Output"));
}
