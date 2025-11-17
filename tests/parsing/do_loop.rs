use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn do_while_loop() {
    let source = r#"
Sub Test()
    Do While x < 10
        x = x + 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("WhileKeyword"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn do_until_loop() {
    let source = r#"
Sub Test()
    Do Until x >= 10
        x = x + 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("UntilKeyword"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn do_loop_while() {
    let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop While x < 10
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("WhileKeyword"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn do_loop_until() {
    let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("UntilKeyword"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn do_loop_infinite() {
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
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn nested_do_loops() {
    let source = r#"
Sub Test()
    Do While i < 10
        Do While j < 5
            j = j + 1
        Loop
        i = i + 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    // Should have two DoStatements
    let do_count = debug.matches("DoStatement").count();
    assert_eq!(do_count, 2);
}

#[test]
fn do_while_with_complex_condition() {
    let source = r#"
Sub Test()
    Do While x < 10 And y > 0
        x = x + 1
        y = y - 1
    Loop
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    assert!(debug.contains("WhileKeyword"));
    assert!(debug.contains("LoopKeyword"));
}

#[test]
fn do_loop_preserves_whitespace() {
    let source = r#"
Sub Test()
    Do  While  x < 10
        x = x + 1
    Loop  While  y > 0
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let debug = cst.debug_tree();
    assert!(debug.contains("DoStatement"));
    // Check that whitespace is preserved
    assert!(debug.contains("Whitespace"));
}
