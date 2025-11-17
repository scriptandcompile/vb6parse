use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn inline_if_then_goto() {
    let source = r#"
Sub Test()
    If x > 0 Then GoTo Positive
    Debug.Print "negative or zero"
Positive:
    Debug.Print "positive"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("ThenKeyword"));
}

#[test]
fn inline_if_then_call() {
    let source = r#"
Sub Test()
    If enabled Then Call DoSomething
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("CallStatement"));
}

#[test]
fn inline_if_then_assignment() {
    let source = r#"
Sub Test()
    If x > 10 Then result = "large"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("AssignmentStatement"));
}

#[test]
fn inline_if_then_set() {
    let source = r#"
Sub Test()
    If obj Is Nothing Then Set obj = New MyClass
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("SetStatement"));
}

#[test]
fn inline_if_then_exit() {
    let source = r#"
Sub Test()
    If errorOccurred Then Exit Sub
    Debug.Print "continuing"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("ExitKeyword"));
}

#[test]
fn inline_if_then_multiple_statements() {
    let source = r#"
Sub Test()
    If condition Then x = 1: y = 2
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    let count = debug.matches("AssignmentStatement").count();
    assert_eq!(count, 2, "Expected 2 assignment statements separated by colon");
}

#[test]
fn inline_if_preserves_whitespace() {
    let source = r#"
Sub Test()
    If x > 0 Then GoTo Label1
Label1:
    x = 1
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("Whitespace"));
    assert!(debug.contains("Newline"));
}

#[test]
fn inline_if_then_goto_with_comment() {
    let source = r#"
Sub Test()
    If x > 0 Then GoTo Positive ' go to positive case
Positive:
    result = x
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("GotoStatement"));
    assert!(debug.contains("EndOfLineComment"));
}

#[test]
fn inline_if_then_call_with_args() {
    let source = r#"
Sub Test()
    If ready Then Call Process(x, y, z)
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("CallStatement"));
}

#[test]
fn inline_if_then_nested_calls() {
    let source = r#"
Sub Test()
    If value > 0 Then result = Calculate(value)
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("AssignmentStatement"));
}

#[test]
fn inline_if_complex_condition() {
    let source = r#"
Sub Test()
    If x > 0 And y < 10 Then GoTo Valid
Valid:
    Process
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("GotoStatement"));
}

#[test]
fn inline_if_not_condition() {
    let source = r#"
Sub Test()
    If Not IsValid Then Exit Sub
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("IfStatement"));
    assert!(debug.contains("ExitKeyword"));
}
