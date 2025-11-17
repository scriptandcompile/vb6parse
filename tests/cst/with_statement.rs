use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

#[test]
fn with_statement_simple() {
    let source = r#"
Sub Test()
    With myObject
        .Property1 = "value"
        .Property2 = 123
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("WithKeyword"));
    assert!(debug.contains("myObject"));
}

#[test]
fn with_statement_nested_property() {
    let source = r#"
Sub Test()
    With obj.SubObject
        .Value = 42
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("SubObject"));
}

#[test]
fn with_statement_method_call() {
    let source = r#"
Sub Test()
    With Form1
        .Show
        .Caption = "Title"
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("Form1"));
}

#[test]
fn with_statement_nested() {
    let source = r#"
Sub Test()
    With outer
        .Value1 = 1
        With .Inner
            .Value2 = 2
        End With
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("WithStatement").count();
    assert_eq!(count, 2, "Expected 2 With statements (nested)");
}

#[test]
fn with_statement_multiple_properties() {
    let source = r#"
Sub Test()
    With employee
        .FirstName = "John"
        .LastName = "Doe"
        .Age = 30
        .Salary = 50000
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("employee"));
}

#[test]
fn with_statement_with_if() {
    let source = r#"
Sub Test()
    With obj
        If .IsValid Then
            .Process
        End If
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("IfStatement"));
}

#[test]
fn with_statement_with_loop() {
    let source = r#"
Sub Test()
    With collection
        For i = 1 To .Count
            .Item(i).Process
        Next i
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("ForStatement"));
}

#[test]
fn with_statement_array_access() {
    let source = r#"
Sub Test()
    With myArray(5)
        .Name = "Test"
        .Value = 100
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("myArray"));
}

#[test]
fn with_statement_function_result() {
    let source = r#"
Sub Test()
    With GetObject()
        .Property = value
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("GetObject"));
}

#[test]
fn with_statement_empty() {
    let source = r#"
Sub Test()
    With obj
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
}

#[test]
fn with_statement_sequential() {
    let source = r#"
Sub Test()
    With obj1
        .Value = 1
    End With
    With obj2
        .Value = 2
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("WithStatement").count();
    assert_eq!(count, 2, "Expected 2 sequential With statements");
}

#[test]
fn with_statement_preserves_whitespace() {
    let source = r#"
With obj
    .Property = value
End With
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn with_statement_new_object() {
    let source = r#"
Sub Test()
    With New MyClass
        .Initialize
        .Value = 42
    End With
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("NewKeyword"));
}

#[test]
fn with_statement_at_module_level() {
    let source = r#"
With GlobalObject
    .Config = value
End With
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("WithStatement"));
    assert!(debug.contains("GlobalObject"));
}
