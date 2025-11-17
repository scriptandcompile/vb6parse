use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn set_statement_simple() {
    let source = r#"
Sub Test()
    Set obj = myObject
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
    assert!(debug.contains("SetKeyword"));
}

#[test]
fn set_statement_with_new() {
    let source = r#"
Sub Test()
    Set obj = New MyClass
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
    assert!(debug.contains("NewKeyword"));
}

#[test]
fn set_statement_to_nothing() {
    let source = r#"
Sub Test()
    Set obj = Nothing
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
}

#[test]
fn set_statement_with_property_access() {
    let source = r#"
Sub Test()
    Set myObj.Property = otherObj
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
    assert!(debug.contains("PeriodOperator"));
}

#[test]
fn set_statement_with_function_call() {
    let source = r#"
Sub Test()
    Set result = GetObject("WinMgmts:")
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
}

#[test]
fn set_statement_with_collection_access() {
    let source = r#"
Sub Test()
    Set item = collection.Item(1)
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
}

#[test]
fn multiple_set_statements() {
    let source = r#"
Sub Test()
    Set obj1 = New Class1
    Set obj2 = New Class2
    Set obj3 = Nothing
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let set_count = debug.matches("SetStatement").count();
    assert_eq!(set_count, 3);
}

#[test]
fn set_statement_preserves_whitespace() {
    let source = r#"
Sub Test()
    Set   obj   =   New   MyClass
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn set_statement_in_function() {
    let source = r#"
Function GetObject() As Object
    Set GetObject = New MyClass
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
    assert!(debug.contains("FunctionStatement"));
}

#[test]
fn set_statement_at_module_level() {
    let source = r#"
Set globalObj = New MyClass
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("SetStatement"));
}
