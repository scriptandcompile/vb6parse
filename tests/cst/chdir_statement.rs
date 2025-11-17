use vb6parse::{parse, tokenize, SourceStream};

#[test]
fn chdir_simple_string_literal() {
    let source = r#"
Sub Test()
    ChDir "C:\Windows"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_variable() {
    let source = r#"
Sub Test()
    ChDir myPath
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_app_path() {
    let source = r#"
Sub Test()
    ChDir App.Path
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_expression() {
    let source = r#"
Sub Test()
    ChDir GetPath() & "\subdir"
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_in_if_statement() {
    let source = r#"
Sub Test()
    If dirExists Then ChDir newPath
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_at_module_level() {
    let source = r#"
ChDir "C:\Temp"
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_comment() {
    let source = r#"
Sub Test()
    ChDir basePath ' Change to base directory
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
    assert!(debug.contains("EndOfLineComment"));
}

#[test]
fn chdir_multiple_in_sequence() {
    let source = r#"
Sub Test()
    ChDir "C:\Windows"
    ChDir "C:\Temp"
    ChDir originalPath
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let chdir_count = debug.matches("ChDirStatement").count();
    assert_eq!(chdir_count, 3, "Expected 3 ChDir statements");
}

#[test]
fn chdir_in_multiline_if() {
    let source = r#"
Sub Test()
    If pathValid Then
        ChDir newPath
    End If
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_parentheses() {
    let source = r#"
Sub Test()
    ChDir (basePath)
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}

#[test]
fn chdir_with_parentheses_without_space() {
    let source = r#"
Sub Test()
    ChDir(basePath)
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("ChDirStatement"));
    assert!(debug.contains("ChDirKeyword"));
}