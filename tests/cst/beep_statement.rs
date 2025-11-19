use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn beep_simple() {
    let source = r#"
Sub Test()
    Beep
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
    assert!(debug.contains("BeepKeyword"));
}

#[test]
fn beep_in_if_statement() {
    let source = r#"
Sub Test()
    If condition Then
        Beep
    End If
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
    assert!(debug.contains("IfStatement"));
}

#[test]
fn beep_inline_if() {
    let source = r#"
Sub Test()
    If error Then Beep
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
}

#[test]
fn multiple_beep_statements() {
    let source = r#"
Sub Test()
    Beep
    Beep
    Beep
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("BeepStatement").count();
    assert_eq!(count, 3);
}

#[test]
fn beep_with_comment() {
    let source = r#"
Sub Test()
    Beep ' Alert user
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
    assert!(debug.contains("EndOfLineComment"));
}

#[test]
fn beep_in_loop() {
    let source = r#"
Sub Test()
    For i = 1 To 3
        Beep
    Next i
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
    assert!(debug.contains("ForStatement"));
}

#[test]
fn beep_in_select_case() {
    let source = r#"
Sub Test()
    Select Case value
        Case 1
            Beep
        Case Else
            Beep
    End Select
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("BeepStatement").count();
    assert_eq!(count, 2);
}

#[test]
fn beep_preserves_whitespace() {
    let source = r#"
Sub Test()
    Beep    ' Extra spaces
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn beep_at_module_level() {
    let source = r#"
Beep
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
}

#[test]
fn beep_with_error_handling() {
    let source = r#"
Sub Test()
    On Error Resume Next
    Beep
    If Err Then MsgBox "Error occurred"
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("BeepStatement"));
}
