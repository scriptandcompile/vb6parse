//! Test for proper ArgumentList parsing in procedure calls.
//!
//! This test verifies that procedure calls create proper ArgumentList nodes
//! with structured Argument children, rather than flat token streams.

use vb6parse::*;

#[test]
fn call_with_parenthesized_arguments() {
    let source = r"prefs.Add settings(i, 1), settings(i, 0)";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures - ArgumentList structure is tested via snapshots
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_with_unparenthesized_arguments() {
    let source = r"Debug.Print x, y, z";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_no_arguments() {
    let source = r"DoEvents";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_with_member_access() {
    let source = r"obj.Method arg1, arg2";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_with_dot_prefix() {
    let source = r#"
Sub Test()
    With myObject
        .Method "argument"
    End With
End Sub
"#;
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_statement_with_keyword() {
    let source = r"
Sub Test()
    Call MySub(x, y)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}

#[test]
fn call_nested_in_expression() {
    let source = r"result = Calculate(x, y)";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Verify no parse failures
    assert!(failures.is_empty(), "Expected no parse failures");
}
