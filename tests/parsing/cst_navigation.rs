use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

// Navigation method tests

#[test]
fn navigation_children() {
    let code = "Attribute VB_Name\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    let children = cst.children();

    assert_eq!(children.len(), 2); // AttributeStatement, SubStatement
    assert_eq!(children[0].kind, SyntaxKind::AttributeStatement);
    assert_eq!(children[1].kind, SyntaxKind::SubStatement);
    assert!(!children[0].is_token);
    assert!(!children[1].is_token);
}

#[test]
fn navigation_find_children_by_kind() {
    let code = "Dim x\nDim y\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Find all DimStatements
    let dim_statements = cst.find_children_by_kind(SyntaxKind::DimStatement);
    assert_eq!(dim_statements.len(), 2);

    // Find all SubStatements
    let sub_statements = cst.find_children_by_kind(SyntaxKind::SubStatement);
    assert_eq!(sub_statements.len(), 1);
}

#[test]
fn navigation_contains_kind() {
    let code = "Sub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert!(cst.contains_kind(SyntaxKind::SubStatement));
    assert!(!cst.contains_kind(SyntaxKind::FunctionStatement));
    assert!(!cst.contains_kind(SyntaxKind::DimStatement));
}

#[test]
fn navigation_first_and_last_child() {
    let code = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let first = cst.first_child().unwrap();
    assert_eq!(first.kind, SyntaxKind::AttributeStatement);
    assert_eq!(first.text, "Attribute VB_Name\n");

    let last = cst.last_child().unwrap();
    assert_eq!(last.kind, SyntaxKind::SubStatement);
}

#[test]
fn navigation_child_at() {
    let code = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let first = cst.child_at(0).unwrap();
    assert_eq!(first.kind, SyntaxKind::AttributeStatement);

    let second = cst.child_at(1).unwrap();
    assert_eq!(second.kind, SyntaxKind::DimStatement);

    let third = cst.child_at(2).unwrap();
    assert_eq!(third.kind, SyntaxKind::SubStatement);

    // Fourth is EOF, out of bounds after that
    assert!(cst.child_at(4).is_none());
}
#[test]
fn navigation_empty_tree() {
    let code = "";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Even empty code has no children now
    assert_eq!(cst.children().len(), 0);
    assert!(cst.first_child().is_none());
    assert!(cst.last_child().is_none());
    assert!(cst.child_at(0).is_none());
    assert!(!cst.contains_kind(SyntaxKind::SubStatement));
}
#[test]
fn navigation_with_comments_and_whitespace() {
    let code = "' Comment\n\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    let children = cst.children();

    // Should have 4 children: EndOfLineComment, newline, newline, SubStatement
    assert_eq!(children.len(), 4);

    // First is the comment
    assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
    assert!(children[0].is_token);

    // Second is newline
    assert_eq!(children[1].kind, SyntaxKind::Newline);
    assert!(children[1].is_token);

    // Third is the second newline
    assert_eq!(children[2].kind, SyntaxKind::Newline);
    assert!(children[2].is_token);

    // Fourth is SubStatement
    assert_eq!(children[3].kind, SyntaxKind::SubStatement);
    assert!(!children[3].is_token);
}
