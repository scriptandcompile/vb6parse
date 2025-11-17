use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;


#[test]
fn cst_distinguishes_declarations_from_functions() {
    // Test that Private declaration and Private Function are correctly distinguished
    let code = "Private myVar As Integer\nPrivate Function GetVar() As Integer\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 2);

    // First child should be a DimStatement (declaration)
    if let Some(first_child) = cst.child_at(0) {
        assert_eq!(first_child.kind, SyntaxKind::DimStatement);
    }

    // Second child should be a FunctionStatement
    if let Some(second_child) = cst.child_at(1) {
        assert_eq!(second_child.kind, SyntaxKind::FunctionStatement);
    }

    assert!(cst.text().contains("Private myVar As Integer"));
    assert!(cst.text().contains("Private Function GetVar"));
}

#[test]
fn cst_all_function_modifier_combinations() {
    // Test all valid function/sub modifier combinations
    let test_cases = vec![
        (
            "Public Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Private Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Friend Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Static Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Public Static Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Private Static Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        (
            "Friend Static Function Test() As Integer\nEnd Function\n",
            SyntaxKind::FunctionStatement,
        ),
        ("Public Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Private Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Friend Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        (
            "Public Static Sub Test()\nEnd Sub\n",
            SyntaxKind::SubStatement,
        ),
        (
            "Private Static Sub Test()\nEnd Sub\n",
            SyntaxKind::SubStatement,
        ),
        (
            "Friend Static Sub Test()\nEnd Sub\n",
            SyntaxKind::SubStatement,
        ),
    ];

    for (code, expected_kind) in test_cases {
        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        assert_eq!(cst.child_count(), 1, "Code: {}", code);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, expected_kind, "Code: {}", code);
        }
    }
}


#[test]
fn cst_function_with_modifiers() {
    // Test Public Function
    let code = "Public Function GetValue() As Integer\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Public Function GetValue"));
}

#[test]
fn cst_private_static_function() {
    // Test Private Static Function
    let code = "Private Static Function Calculate(x As Long) As Long\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Private Static Function Calculate"));
}

#[test]
fn cst_friend_function() {
    // Test Friend Function
    let code = "Friend Function ProcessData() As String\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Friend Function ProcessData"));
}