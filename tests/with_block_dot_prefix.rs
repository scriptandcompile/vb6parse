//! Test for With block dot-prefixed member access parsing.
//!
//! This test verifies that dot-prefixed member access in With blocks
//! (like `.Property = value` and `.Method arg`) is correctly parsed
//! without Unknown tokens.
//!
//! Previously, when using `.Property` inside an If condition within a With block,
//! the period operator and property name were parsed incorrectly, and "End If"
//! was parsed as an Unknown "End" token. These tests verify the fixes.

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

/// Test that With block member access in If conditions is parsed correctly.
///
/// Previously, when using `.Property` inside an If condition within a With block,
/// the period operator and property name were parsed incorrectly, and "End If"
/// was parsed as an Unknown "End" token.
#[test]
fn test_with_block_if_condition_member_access() {
    let source = r"
Sub Test()
    With myObject
        If .Property > 0 Then
            .Method
        End If
    End With
End Sub
";

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    // Verify that there are no Unknown tokens
    assert!(
        !tree_str.contains("Unknown"),
        "CST should not contain any Unknown tokens for With block with If statement"
    );

    // Verify that the member access is properly structured
    assert!(
        tree_str.contains("IdentifierExpression"),
        "CST should contain IdentifierExpression for .Property"
    );
    assert!(
        tree_str.contains("BinaryExpression"),
        "CST should contain BinaryExpression for .Property > 0"
    );
}

/// Test that With block member access in various contexts is parsed correctly.
#[test]
fn test_with_block_member_access_contexts() {
    let source = r#"
Sub Test()
    With myObject
        .StringProp = "value"
        
        If .GetValue() > 10 Then
            .DoSomething
        End If
        
        If .Count = 0 Then
            MsgBox "Empty"
        End If
    End With
End Sub
"#;

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    // Verify that there are no Unknown tokens
    assert!(
        !tree_str.contains("Unknown"),
        "CST should not contain any Unknown tokens for various With block member accesses"
    );
}

/// Test With block with If checking property - ensures no Unknown tokens.
#[test]
fn test_with_block_property_check() {
    let with_source = r"
Sub Test()
    With myObject
        If .Active = True Then
            .DoSomething
        End If
    End With
End Sub
";

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", with_source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    assert!(
        !tree_str.contains("Unknown"),
        "With block with If should not have Unknown tokens"
    );
}
