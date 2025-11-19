use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn parse_option_explicit_on() {
    let code = "Option Explicit On\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Option statements are currently parsed as DimStatement (declarations)
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit On\n");
}

#[test]
fn parse_option_explicit_off() {
    let code = "Option Explicit Off\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit Off\n");
}

#[test]
fn parse_option_explicit() {
    let code = "Option Explicit\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit\n");
}
