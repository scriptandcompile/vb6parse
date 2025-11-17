use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn exit_do() {
    let source = r#"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("DoKeyword"));
}

#[test]
fn exit_for() {
    let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 5 Then Exit For
    Next
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("ForKeyword"));
}

#[test]
fn exit_function() {
    let source = r#"
Function Test() As Integer
    If x = 0 Then
        Exit Function
    End If
    Test = 42
End Function
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("FunctionKeyword"));
}

#[test]
fn exit_sub() {
    let source = r#"
Sub Test()
    If x = 0 Then Exit Sub
    Debug.Print "x is not zero"
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("SubKeyword"));
}

#[test]
fn exit_property() {
    // TODO: Fix this to a Property procedure when supported
    let source = r#"
Sub TestPropertyStub()
    If m_value = "" Then Exit Property
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("PropertyKeyword"));
}

#[test]
fn multiple_exit_statements() {
    let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 3 Then Exit For
        If i = 7 Then Exit Sub
    Next
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    // Should have two ExitStatements
    let exit_count = debug.matches("ExitStatement").count();
    assert_eq!(exit_count, 2);
}

#[test]
fn exit_in_nested_loops() {
    let source = r#"
Sub Test()
    Do While x < 100
        For i = 1 To 10
            If i = 5 Then Exit For
        Next
        If x > 50 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    let exit_count = debug.matches("ExitStatement").count();
    assert_eq!(exit_count, 2);
}

#[test]
fn exit_preserves_whitespace() {
    let source = r#"
Sub Test()
    Exit   Sub
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("ExitStatement"));
    // Check that whitespace is preserved
    assert!(debug.contains("Whitespace"));
}

#[test]
fn inline_exit_in_if_statement() {
    let source = r#"
Function Test(x As Integer) As Integer
    If x = 0 Then Exit Function
    Test = x * 2
End Function
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("ExitStatement"));
    assert!(debug.contains("ExitKeyword"));
    assert!(debug.contains("FunctionKeyword"));
}
