use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn parse_dim_declaration() {
    let code = "Dim x As Integer\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Dim x As Integer\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
}

#[test]
fn parse_private_declaration() {
    let code = "Private m_value As Long\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Private m_value As Long\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PrivateKeyword"));
}

#[test]
fn parse_public_declaration() {
    let code = "Public g_config As String\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Public g_config As String\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PublicKeyword"));
}

#[test]
fn parse_multiple_variable_declaration() {
    let code = "Dim x, y, z As Integer\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Dim x, y, z As Integer\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
}

#[test]
fn parse_const_declaration() {
    let code = "Const MAX_SIZE = 100\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Const MAX_SIZE = 100\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("ConstKeyword"));
}

#[test]
fn parse_private_const_declaration() {
    let code = "Private Const MODULE_NAME = \"MyModule\"\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Private Const MODULE_NAME = \"MyModule\"\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PrivateKeyword"));
    assert!(debug.contains("ConstKeyword"));
}

#[test]
fn parse_static_declaration() {
    let code = "Static counter As Long\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Static counter As Long\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("StaticKeyword"));
}
