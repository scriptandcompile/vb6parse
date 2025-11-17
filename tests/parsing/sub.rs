use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;


#[test]
fn cst_public_static_sub() {
    // Test Public Static Sub
    let code = "Public Static Sub Initialize()\nEnd Sub\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::SubStatement);
    }
    assert!(cst.text().contains("Public Static Sub Initialize"));
}


#[test]
fn cst_public_sub() {
    // Test Public Sub
    let code = "Public Sub Initialize()\nEnd Sub\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::SubStatement);
    }
    assert!(cst.text().contains("Public Sub Initialize"));
}


#[test]
fn cst_private_sub() {
    // Test Private Sub
    let code = "Private Sub Initialize()\nEnd Sub\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::SubStatement);
    }
    assert!(cst.text().contains("Private Sub Initialize"));
}