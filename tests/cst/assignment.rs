use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

fn create_cst_from_source(source: &str) -> vb6parse::parsers::cst::ConcreteSyntaxTree {
    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    parse(token_stream)
}

#[test]
fn test_simple_assignment() {
    let source = r#"
x = 5
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "x");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[4].text, "5");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);
    
    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_string_assignment() {
    let source = r#"
myName = "John"
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "myName");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::StringLiteral);
    assert_eq!(cst.children()[1].children[4].text, "\"John\"");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Newline);
    
    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_property_assignment() {
    let source = r#"
obj.subProperty = value
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    // The assignment contains: obj.subProperty = value
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "obj");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[2].text, "subProperty");
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[6].text, "value");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_array_assignment() {
    let source = r#"
arr(0) = 100
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "arr");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[2].text, "0");
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[7].text, "100");
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_multidimensional_array_assignment() {
    let source = r#"
matrix(i, j) = value
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "matrix");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[2].text, "i");
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Comma);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[5].text, "j");
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[10].text, "value");
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_function_call() {
    let source = r#"
result = MyFunction(arg1, arg2)
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "result");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "MyFunction");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[6].text, "arg1");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Comma);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[9].text, "arg2");
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_expression() {
    let source = r#"
sum = a + b * c
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "sum");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "a");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::AdditionOperator);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[8].text, "b");
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::MultiplicationOperator);
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[12].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[12].text, "c");
    assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Newline);


    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_method_call() {
    let source = r#"
text = obj.GetText()
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "text");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "obj");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[6].text, "GetText");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_nested_property() {
    let source = r#"
value = obj.SubObj.SubProperty
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "value");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "obj");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[6].text, "SubObj");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[8].text, "SubProperty");
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_multiple_assignments() {
    let source = r#"
x = 1
y = 2
z = 3
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "x");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[4].text, "1");

    assert_eq!(cst.children()[2].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[2].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[2].children[0].text, "y");
    assert_eq!(cst.children()[2].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[2].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[2].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[2].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[2].children[4].text, "2");

    assert_eq!(cst.children()[3].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[3].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[3].children[0].text, "z");
    assert_eq!(cst.children()[3].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[3].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[3].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[3].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[3].children[4].text, "3");

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_preserves_whitespace() {
    let source = "x   =   5";
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[0].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[0].children[0].text, "x");
    assert_eq!(cst.children()[0].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[0].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[0].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[0].children[4].kind, SyntaxKind::Number);
    
    // Verify whitespace is preserved
    assert_eq!(cst.text(), source);
}

#[test]
fn test_assignment_in_function() {
    let source = r#"
Public Function Calculate()
    result = 42
End Function
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::FunctionStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::PublicKeyword);
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::FunctionKeyword);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "Calculate");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::ParameterList);
    assert_eq!(cst.children()[1].children[5].children[0].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[5].children[1].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Newline);

    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::CodeBlock);
    assert_eq!(cst.children()[1].children[7].children[0].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[7].children[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[7].children[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[7].children[1].children[0].text, "result");
    assert_eq!(cst.children()[1].children[7].children[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[7].children[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[7].children[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[7].children[1].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[7].children[1].children[4].text, "42");
    assert_eq!(cst.children()[1].children[7].children[1].children[5].kind, SyntaxKind::Newline);

    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::EndKeyword);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::FunctionKeyword);
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_collection_access() {
    let source = r#"
item = Collection("Key")
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "item");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "Collection");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::StringLiteral);
    assert_eq!(cst.children()[1].children[6].text, "\"Key\"");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_dollar_sign_function() {
    let source = r#"
path = Environ$("TEMP")
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "path");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "Environ");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::DollarSign);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::StringLiteral);
    assert_eq!(cst.children()[1].children[7].text, "\"TEMP\"");
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_at_module_level() {
    let source = r#"
Option Explicit
x = 5
"#;
    let cst = create_cst_from_source(source);
    
    assert!(cst.children()[0].kind == SyntaxKind::Newline);

    assert!(cst.children()[1].kind == SyntaxKind::OptionStatement);
    assert!(cst.children()[1].children[0].kind == SyntaxKind::OptionKeyword);
    assert!(cst.children()[1].children[1].kind == SyntaxKind::Whitespace);
    assert!(cst.children()[1].children[2].kind == SyntaxKind::ExplicitKeyword);
    assert!(cst.children()[1].children[3].kind == SyntaxKind::Newline);

    assert!(cst.children()[2].kind == SyntaxKind::AssignmentStatement);
    assert!(cst.children()[2].children[0].kind == SyntaxKind::Identifier);
    assert!(cst.children()[2].children[0].text == "x");
    assert!(cst.children()[2].children[1].kind == SyntaxKind::Whitespace);
    assert!(cst.children()[2].children[2].kind == SyntaxKind::EqualityOperator);
    assert!(cst.children()[2].children[3].kind == SyntaxKind::Whitespace);
    assert!(cst.children()[2].children[4].kind == SyntaxKind::Number);
    assert!(cst.children()[2].children[4].text == "5");
    assert!(cst.children()[2].children[5].kind == SyntaxKind::Newline);
    
    // Verify the parsed tree can be converted back to the original source
    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_numeric_literal() {
    let source = r#"
pi = 3.14159
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "pi");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[4].text, "3");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[6].text, "14159");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_concatenation() {
    let source = r#"
fullName = firstName & " " & lastName
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "fullName");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[4].text, "firstName");
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Ampersand);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::StringLiteral);
    assert_eq!(cst.children()[1].children[8].text, "\" \"");
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::Ampersand);
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[12].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[12].text, "lastName");
    assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Newline);
    
    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_to_type_member() {
    let source = r#"
person.Age = 25
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "person");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::PeriodOperator);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[2].text, "Age");
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Number);
    assert_eq!(cst.children()[1].children[6].text, "25");
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}

#[test]
fn test_assignment_with_parenthesized_expression() {
    let source = r#"
result = (a + b) * c
"#;
    let cst = create_cst_from_source(source);

    assert_eq!(cst.children()[0].kind, SyntaxKind::Newline);
    assert_eq!(cst.children()[1].kind, SyntaxKind::AssignmentStatement);
    assert_eq!(cst.children()[1].children[0].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[0].text, "result");
    assert_eq!(cst.children()[1].children[1].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[2].kind, SyntaxKind::EqualityOperator);
    assert_eq!(cst.children()[1].children[3].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[4].kind, SyntaxKind::LeftParenthesis);
    assert_eq!(cst.children()[1].children[5].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[5].text, "a");
    assert_eq!(cst.children()[1].children[6].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[7].kind, SyntaxKind::AdditionOperator);
    assert_eq!(cst.children()[1].children[8].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[9].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[9].text, "b");
    assert_eq!(cst.children()[1].children[10].kind, SyntaxKind::RightParenthesis);
    assert_eq!(cst.children()[1].children[11].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[12].kind, SyntaxKind::MultiplicationOperator);
    assert_eq!(cst.children()[1].children[13].kind, SyntaxKind::Whitespace);
    assert_eq!(cst.children()[1].children[14].kind, SyntaxKind::Identifier);
    assert_eq!(cst.children()[1].children[14].text, "c");
    assert_eq!(cst.children()[1].children[15].kind, SyntaxKind::Newline);

    assert_eq!(cst.text().trim(), source.trim());
}
