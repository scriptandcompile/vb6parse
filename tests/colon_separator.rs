//! Tests for VB6 colon (`:`) statement separator.
//!
//! In VB6, a colon can be used to separate multiple statements on the same line:
//! `a = 1 : b = 2 : c = 3`
//!
//! This is distinct from a **label**, where an identifier is followed by a colon
//! at the start of a statement (e.g. `ErrorHandler:`).
//!
//! These tests verify that the colon is parsed as a `ColonOperator` token (not
//! `Unknown`) and that the statements on either side are fully parsed.

use vb6parse::*;

/// Two assignments on the same line separated by a colon.
#[test]
fn two_assignments_colon_separated() {
    let source = r"
Sub Test()
    a = 1 : b = 2
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );
    assert!(
        text.contains("ColonOperator"),
        "Should contain ColonOperator token"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Three assignments on the same line separated by colons.
#[test]
fn three_assignments_colon_separated() {
    let source = r"
Sub Test()
    a = 0 : b = 1 : c = -1
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Procedure calls separated by colons.
#[test]
fn procedure_calls_colon_separated() {
    let source = r"
Sub Test()
    Beep : Beep : Beep
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Mixed statement types separated by colons.
#[test]
fn mixed_statements_colon_separated() {
    let source = r"
Sub Test()
    x = 10 : Print x : y = x + 1
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator with no space around it.
#[test]
fn colon_separator_no_spaces() {
    let source = r"
Sub Test()
    a = 1:b = 2
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator inside an If-Then body.
#[test]
fn colon_separator_in_if_body() {
    let source = r"
Sub Test()
    If x > 0 Then
        a = 1 : b = 2
    End If
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator inside a For loop body.
#[test]
fn colon_separator_in_for_loop() {
    let source = r"
Sub Test()
    For i = 1 To 10
        a = i : b = i * 2
    Next i
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator inside a Do-Loop body.
#[test]
fn colon_separator_in_do_loop() {
    let source = r"
Sub Test()
    Do While i < 10
        i = i + 1 : total = total + i
    Loop
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator inside a With block.
#[test]
fn colon_separator_in_with_block() {
    let source = r"
Sub Test()
    With Triangle
        .Width = 100 : .Height = 200
    End With
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator with a comment at end of line.
#[test]
fn colon_separator_with_trailing_comment() {
    let source = r"
Sub Test()
    a = 1 : b = 2 ' two assignments
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Label (identifier + colon) is NOT a statement separator — it must still parse correctly.
/// This test ensures we don't regress label parsing.
#[test]
fn label_still_parses_correctly() {
    let source = r"
Sub Test()
    GoTo MyLabel
MyLabel:
    a = 1
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );
    assert!(
        text.contains("LabelStatement"),
        "Should contain LabelStatement"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Dim statement followed by assignment via colon separator.
#[test]
fn dim_then_assignment_colon_separated() {
    let source = r"
Sub Test()
    Dim x As Integer : x = 42
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens, found:\n{text}"
    );

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

/// Colon separator in a single-line If-Then statement.
#[test]
fn colon_separator_in_single_line_if() {
    let source = r"
Sub Test()
    If x > 0 Then a = 1 : b = 2
End Sub
";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/colon_separator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}
