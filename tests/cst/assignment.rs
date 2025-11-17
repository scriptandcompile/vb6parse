use vb6parse::parsers::{parse, SourceStream};
use vb6parse::tokenize::tokenize;

fn test_assignment(source: &str) {
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AssignmentStatement"), "No assignment statement found");
}

#[test]
fn test_simple_assignment() {
    let source = r#"
x = 5
"#;
    test_assignment(source);
}

#[test]
fn test_string_assignment() {
    let source = r#"
name = "John"
"#;
    test_assignment(source);
}

#[test]
fn test_property_assignment() {
    let source = r#"
obj.Property = value
"#;
    test_assignment(source);
}

#[test]
fn test_array_assignment() {
    let source = r#"
arr(0) = 100
"#;
    test_assignment(source);
}

#[test]
fn test_multidimensional_array_assignment() {
    let source = r#"
matrix(i, j) = value
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_function_call() {
    let source = r#"
result = MyFunction(arg1, arg2)
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_expression() {
    let source = r#"
sum = a + b * c
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_method_call() {
    let source = r#"
text = obj.GetText()
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_nested_property() {
    let source = r#"
value = obj.SubObj.Property
"#;
    test_assignment(source);
}

#[test]
fn test_multiple_assignments() {
    let source = r#"
x = 1
y = 2
z = 3
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    let count = debug.matches("AssignmentStatement").count();
    assert_eq!(count, 3, "Expected 3 assignment statements");
}

#[test]
fn test_assignment_preserves_whitespace() {
    let source = "x   =   5";
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AssignmentStatement"));
    assert!(debug.contains("Whitespace"));
}

#[test]
fn test_assignment_in_function() {
    let source = r#"
Public Function Calculate()
    result = 42
End Function
"#;
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    assert!(debug.contains("AssignmentStatement"));
    assert!(debug.contains("FunctionStatement"));
}

#[test]
fn test_assignment_with_collection_access() {
    let source = r#"
item = Collection("Key")
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_dollar_sign_function() {
    let source = r#"
path = Environ$("TEMP")
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_at_module_level() {
    let source = r#"
Option Explicit
x = 5
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_numeric_literal() {
    let source = r#"
pi = 3.14159
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_concatenation() {
    let source = r#"
fullName = firstName & " " & lastName
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_to_type_member() {
    let source = r#"
person.Age = 25
"#;
    test_assignment(source);
}

#[test]
fn test_assignment_with_parenthesized_expression() {
    let source = r#"
result = (a + b) * c
"#;
    test_assignment(source);
}
