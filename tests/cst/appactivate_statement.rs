use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn appactivate_simple() {
    let source = r#"
Sub Test()
    AppActivate "MyApp"
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
    assert!(debug.contains("AppActivateKeyword"));
}

#[test]
fn appactivate_with_variable() {
    let source = r#"
Sub Test()
    AppActivate lstTopWin.Text
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}

#[test]
fn appactivate_with_wait_parameter() {
    let source = r#"
Sub Test()
    AppActivate "Calculator", True
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}

#[test]
fn appactivate_with_title_variable() {
    let source = r#"
Sub Test()
    AppActivate sTitle
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}

#[test]
fn appactivate_preserves_whitespace() {
    let source = r#"
Sub Test()
    AppActivate   "MyApp"  ,  False
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn multiple_appactivate_statements() {
    let source = r#"
Sub Test()
    AppActivate "App1"
    AppActivate "App2"
    AppActivate windowTitle
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("AppActivateStatement").count();
    assert_eq!(count, 3);
}

#[test]
fn appactivate_in_if_statement() {
    let source = r#"
Sub Test()
    If condition Then
        AppActivate "MyApp"
    End If
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
    assert!(debug.contains("IfStatement"));
}

#[test]
fn appactivate_inline_if() {
    let source = r#"
Sub Test()
    If windowExists Then AppActivate windowTitle
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}

#[test]
fn appactivate_with_error_handling() {
    let source = r#"
Sub Test()
    On Error Resume Next
    AppActivate lstTopWin.Text
    If Err Then MsgBox "AppActivate error: " & Err
End Sub
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}

#[test]
fn appactivate_at_module_level() {
    let source = r#"
AppActivate "MyApp"
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AppActivateStatement"));
}
