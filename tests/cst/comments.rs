//! Integration test for CST parsing functionality revolving around comments.

use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn parse_single_quote_comment() {
    let code = "' This is a comment\nSub Main()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 2 children: the comment and the SubStatement
    assert_eq!(cst.child_count(), 3); // 2 statements + EOF
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
    assert!(cst.contains_kind(SyntaxKind::SubStatement));

    let first = cst.first_child().unwrap();
    assert_eq!(first.kind, SyntaxKind::EndOfLineComment);
    assert!(first.is_token);
}

#[test]
fn parse_rem_comment() {
    let code = "REM This is a REM comment\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 2 children: the REM comment and the SubStatement
    assert_eq!(cst.child_count(), 3); // 2 statements + EOF
    assert!(cst.text().contains("REM This is a REM comment"));
    assert!(cst.text().contains("Sub Test()"));

    // Verify REM comment is preserved
    let debug = cst.debug_tree();
    assert!(debug.contains("RemComment"));
}

#[test]
fn parse_mixed_comments() {
    let code = "' Single quote comment\nREM REM comment\nSub Test()\nEnd Sub\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 5 children: EndOfLineComment, Newline, RemComment, Newline, SubStatement
    assert_eq!(cst.child_count(), 5);
    assert!(cst.text().contains("' Single quote comment"));
    assert!(cst.text().contains("REM REM comment"));

    // Use navigation methods
    let children = cst.children();
    assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
    assert_eq!(children[1].kind, SyntaxKind::Newline);
    assert_eq!(children[2].kind, SyntaxKind::RemComment);
    assert_eq!(children[3].kind, SyntaxKind::Newline);
    assert_eq!(children[4].kind, SyntaxKind::SubStatement);

    assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
    assert!(cst.contains_kind(SyntaxKind::RemComment));
}

#[test]
fn cst_with_comments() {
    let code = "' This is a comment\nSub Main()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Now has 3 children: comment token, newline token, SubStatement
    assert_eq!(cst.child_count(), 3);
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));
}
