use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn simple_for_loop() {
    let source = r#"
Sub TestSub()
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("ForKeyword"));
    assert!(debug.contains("ToKeyword"));
    assert!(debug.contains("NextKeyword"));
}

#[test]
fn for_loop_with_step() {
    let source = r#"
Sub TestSub()
    For i = 1 To 100 Step 5
        Debug.Print i
    Next i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("StepKeyword"));
}

#[test]
fn for_loop_with_negative_step() {
    let source = r#"
Sub TestSub()
    For i = 10 To 1 Step -1
        Debug.Print i
    Next i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("StepKeyword"));
}

#[test]
fn for_loop_without_counter_after_next() {
    let source = r#"
Sub TestSub()
    For i = 1 To 10
        Debug.Print i
    Next
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("NextKeyword"));
}

#[test]
fn nested_for_loops() {
    let source = r#"
Sub TestSub()
    For i = 1 To 5
        For j = 1 To 5
            Debug.Print i * j
        Next j
    Next i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    // Count occurrences of ForStatement - should have 2
    let for_count = debug.matches("ForStatement").count();
    assert_eq!(for_count, 2);
}

#[test]
fn for_loop_with_function_calls() {
    let source = r#"
Sub TestSub()
    For i = GetStart() To GetEnd() Step GetStep()
        Debug.Print i
    Next i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("ToKeyword"));
    assert!(debug.contains("StepKeyword"));
}

#[test]
fn for_loop_preserves_whitespace() {
    let source = r#"
Sub TestSub()
    For   i   =   1   To   10   Step   2
        Debug.Print i
    Next   i
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ForStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn multiple_for_loops_in_sequence() {
    let source = r#"
Sub TestSub()
    For i = 1 To 5
        Debug.Print "First: " & i
    Next i
    
    For j = 10 To 20 Step 2
        Debug.Print "Second: " & j
    Next j
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    // Count occurrences of ForStatement - should have 2
    let for_count = debug.matches("ForStatement").count();
    assert_eq!(for_count, 2);
}
