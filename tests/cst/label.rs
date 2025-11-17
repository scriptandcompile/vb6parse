use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn label_simple() {
    let source = r#"
Sub Test()
    MyLabel:
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("MyLabel"));
}

#[test]
fn label_with_goto() {
    let source = r#"
Sub Test()
    GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    MsgBox "Error"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("ErrorHandler"));
}

#[test]
fn label_with_underscore() {
    let source = r#"
Sub Test()
Error_Handler:
    MsgBox "Error"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("Error_Handler"));
}

#[test]
fn label_at_module_level() {
    let source = r#"
Sub Test()
StartHere:
    x = 1
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("StartHere"));
}

#[test]
fn label_multiple() {
    let source = r#"
Sub Test()
Start:
    x = 1
Middle:
    y = 2
End_Label:
    z = 3
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("LabelStatement").count();
    assert_eq!(count, 3, "Expected 3 label statements");
}

#[test]
fn label_with_space_after_colon() {
    let source = r#"
Sub Test()
MyLabel: x = 1
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("MyLabel"));
}

#[test]
fn label_error_handler_pattern() {
    let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    ' Some code
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("ErrorHandler"));
}

#[test]
fn label_with_numbers() {
    let source = r#"
Sub Test()
Label123:
    x = 1
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("Label123"));
}

#[test]
fn label_cleanup_pattern() {
    let source = r#"
Sub Test()
    GoTo Cleanup
Cleanup:
    Set obj = Nothing
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("Cleanup"));
}

#[test]
fn label_preserves_whitespace() {
    let source = "MyLabel:";
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("MyLabel"));
    assert!(debug.contains("ColonOperator"));
}

#[test]
fn label_in_function() {
    let source = r#"
Function Calculate() As Integer
Start:
    Calculate = 42
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("Start"));
    assert!(debug.contains("FunctionStatement"));
}

#[test]
fn label_mixed_case() {
    let source = r#"
Sub Test()
MyErrorHandler:
    MsgBox "Error"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("LabelStatement"));
    assert!(debug.contains("MyErrorHandler"));
}
