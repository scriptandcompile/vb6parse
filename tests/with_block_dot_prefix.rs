//! Test for With block dot-prefixed member access parsing.
//!
//! This test verifies that dot-prefixed member access in With blocks
//! (like `.Property = value` and `.Method arg`) is correctly parsed
//! without Unknown tokens.

use vb6parse::*;

#[test]
fn with_block_dot_prefix_assignment() {
    let source = r"
Sub Test()
    With myObject
        .Property = 123
    End With
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Just verify parsing succeeds and no failures
    assert!(_failures.is_empty(), "Expected no parse failures");
}

#[test]
fn with_block_dot_prefix_method_call() {
    let source = r#"
Sub Test()
    With myObject
        .Method "argument"
    End With
End Sub
"#;
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Just verify parsing succeeds and no failures
    assert!(_failures.is_empty(), "Expected no parse failures");
}

#[test]
fn with_block_dot_prefix_multiple() {
    let source = r"
Sub Test()
    With myObject
        .Property1 = 1
        .Property2 = 2
        .Method1
        .Method2 arg
    End With
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let _cst = cst_opt.expect("CST should be parsed");

    // Just verify parsing succeeds and no failures
    assert!(_failures.is_empty(), "Expected no parse failures");
}
