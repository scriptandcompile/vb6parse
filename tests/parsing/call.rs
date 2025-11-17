use vb6parse::parsers::{parse, SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn call_statement_simple() {
    // Test a simple Call statement
    let code = "Call MySubroutine()\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::CallStatement);
    }
    
    assert_eq!(cst.text(), code);
}

#[test]
fn call_statement_with_arguments() {
    // Test a Call statement with arguments
    let code = "Call ProcessData(x, y, z)\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::CallStatement);
    }
    
    assert!(cst.text().contains("Call ProcessData"));
    assert!(cst.text().contains("x, y, z"));
}

#[test]
fn call_statement_preserves_whitespace() {
    // Test that Call statement preserves all whitespace
    let code = "Call  MyFunction (  arg1 ,  arg2  )\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // CST should preserve exact text including all whitespace
    assert_eq!(cst.text(), code);
}

#[test]
fn call_statement_in_sub() {
    // Test Call statement inside a Sub
    let code = "Sub Main()\nCall DoSomething()\nEnd Sub\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(sub_statement) = cst.child_at(0) {
        assert_eq!(sub_statement.kind, SyntaxKind::SubStatement);
        
        // The SubStatement should contain a CallStatement in its children
        assert!(sub_statement.text.contains("Call DoSomething"));
    }
    
    assert_eq!(cst.text(), code);
}

#[test]
fn call_statement_no_parentheses() {
    // Test Call statement without parentheses (valid in VB6)
    let code = "Call MySubroutine\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::CallStatement);
    }
    
    assert_eq!(cst.text(), code);
}

#[test]
fn multiple_call_statements() {
    // Test multiple Call statements
    let code = "Call First()\nCall Second()\nCall Third()\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 3);
    
    for i in 0..3 {
        if let Some(child) = cst.child_at(i) {
            assert_eq!(child.kind, SyntaxKind::CallStatement);
        }
    }
    
    assert!(cst.text().contains("Call First"));
    assert!(cst.text().contains("Call Second"));
    assert!(cst.text().contains("Call Third"));
}

#[test]
fn call_statement_with_string_arguments() {
    // Test Call statement with string literal arguments
    let code = "Call ShowMessage(\"Hello, World!\")\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::CallStatement);
    }
    
    assert!(cst.text().contains("\"Hello, World!\""));
}

#[test]
fn call_statement_with_complex_expressions() {
    // Test Call statement with complex expressions as arguments
    let code = "Call Calculate(x + y, z * 2, (a - b) / c)\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.child_count(), 1);
    
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::CallStatement);
    }
    
    assert!(cst.text().contains("x + y"));
    assert!(cst.text().contains("z * 2"));
}
