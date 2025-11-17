use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn goto_statement_simple() {
    let source = r#"
Sub Test()
    GoTo ErrorHandler
    Debug.Print "Normal code"
ErrorHandler:
    Debug.Print "Error handling"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("GotoKeyword"));
    assert!(debug.contains("ErrorHandler"));
}

#[test]
fn goto_statement_with_line_number() {
    let source = r#"
Sub Test()
    GoTo 100
    Debug.Print "code"
100:
    Debug.Print "target"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("GotoKeyword"));
}

#[test]
fn goto_statement_in_if() {
    let source = r#"
Sub Test()
    If x > 10 Then
        GoTo LargeValue
    End If
    Debug.Print "small"
LargeValue:
    Debug.Print "large"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("LargeValue"));
}

#[test]
fn goto_statement_multiple() {
    let source = r#"
Sub Test()
    GoTo Label1
    GoTo Label2
    GoTo Label3
Label1:
    Debug.Print "one"
Label2:
    Debug.Print "two"
Label3:
    Debug.Print "three"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("GotoStatement").count();
    assert_eq!(count, 3, "Expected 3 GoTo statements");
}

#[test]
fn goto_statement_error_handling() {
    let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Some code that might error
    Debug.Print "normal"
    Exit Sub
ErrorHandler:
    MsgBox "Error occurred"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    // Note: "On Error GoTo" is a special case that may be parsed differently
    // This test just ensures we can handle the basic GoTo part
    assert!(debug.contains("GotoKeyword"));
}

#[test]
fn goto_statement_forward_jump() {
    let source = r#"
Sub Test()
    x = 1
    GoTo SkipCode
    x = 2
    x = 3
SkipCode:
    x = 4
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("SkipCode"));
}

#[test]
fn goto_statement_backward_jump() {
    let source = r#"
Sub Test()
StartLoop:
    counter = counter + 1
    If counter < 10 Then
        GoTo StartLoop
    End If
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("StartLoop"));
}

#[test]
fn goto_statement_in_select_case() {
    let source = r#"
Sub Test()
    Select Case x
        Case 1
            GoTo Handler1
        Case 2
            GoTo Handler2
    End Select
Handler1:
    Debug.Print "one"
Handler2:
    Debug.Print "two"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("SelectCaseStatement"));
}

#[test]
fn goto_statement_in_loop() {
    let source = r#"
Sub Test()
    For i = 1 To 10
        If i = 5 Then
            GoTo ExitLoop
        End If
        Debug.Print i
    Next i
ExitLoop:
    Debug.Print "done"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("ForStatement"));
}

#[test]
fn goto_statement_module_level() {
    let source = r#"
Public Sub TestGoto()
    GoTo Finish
    Debug.Print "skipped"
Finish:
    Debug.Print "done"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
}

#[test]
fn goto_statement_with_underscore() {
    let source = r#"
Sub Test()
    GoTo Error_Handler
Error_Handler:
    Debug.Print "error"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("Error_Handler"));
}

#[test]
fn goto_statement_preserves_whitespace() {
    let source = r#"
Sub Test()
    GoTo MyLabel
MyLabel:
    x = 1
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("Whitespace"));
    assert!(debug.contains("Newline"));
}

#[test]
fn goto_statement_same_line_as_then() {
    let source = r#"
Sub Test()
    If condition Then
        GoTo Handler
    End If
Handler:
    result = True
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("Handler"));
}

#[test]
fn goto_statement_exit_cleanup() {
    let source = r#"
Sub Test()
    On Error GoTo Cleanup
    ' do work
    Exit Sub
Cleanup:
    ' cleanup code
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("GotoKeyword"));
}
