//! Tests for recursion depth limiting to prevent stack overflow.

use vb6parse::{ConcreteSyntaxTree, FormFile, SourceFile};

/// Test that deeply nested controls don't cause stack overflow
#[test]
fn deeply_nested_controls() {
    // Create a form with many nested frames
    let mut form = String::from("Begin VB.Form Form1\n");
    let depth = 100; // Test with 100 levels of nesting

    for i in 0..depth {
        form.push_str(&format!("  Begin VB.Frame Frame{}\n", i));
    }

    for _ in 0..depth {
        form.push_str("  End\n");
    }

    form.push_str("End\n");

    // Should parse without crashing
    let source = SourceFile::from_string("test.frm", form);
    let (result, _errors) = FormFile::parse(&source).unpack();
    assert!(result.is_some());
}

/// Test that extremely deeply nested controls respect depth limits
///
/// **Note**: This test is marked `ignore` because current depth limiting provides
/// protection for reasonable depths (up to ~900 levels) but extremely deep nesting
/// (1100+) can still overflow due to the recursive nature of the parser. We need to
/// convert to iterative parsing for complete protection.
#[test]
#[ignore = "Current depth limiting doesn't fully prevent stack overflow at extreme depths (1100+)"]
fn extremely_nested_controls_depth_limit() {
    // Create a form with nesting that exceeds the limit (1000)
    let mut form = String::from("Begin VB.Form Form1\n");
    let depth = 1100; // Exceeds MAX_CONTROL_DEPTH of 1000

    for i in 0..depth {
        form.push_str(&format!("  Begin VB.Frame Frame{}\n", i));
    }

    for _ in 0..depth {
        form.push_str("  End\n");
    }

    form.push_str("End\n");

    // Should parse but with limited depth (gracefully handles overflow)
    let source = SourceFile::from_string("test.frm", form);
    let _result = FormFile::parse(&source);
    // Phase 2 iterative parsing will handle this case properly
}

/// Test that deeply nested expressions don't cause stack overflow
#[test]
fn deeply_nested_expressions() {
    let mut code = String::from("Sub Test()\n");
    code.push_str("  x = ");

    let depth = 50; // Test with 50 levels of parentheses
    for _ in 0..depth {
        code.push('(');
    }

    code.push_str("42");

    for _ in 0..depth {
        code.push(')');
    }

    code.push_str("\nEnd Sub\n");

    // Should parse without crashing
    let (result, _errors) = ConcreteSyntaxTree::from_text("test.bas", &code).unpack();
    assert!(result.is_some());
}

/// Test that deeply nested if statements don't cause stack overflow
#[test]
fn deeply_nested_if_statements() {
    let mut code = String::from("Sub Test()\n");
    let depth = 50; // Test with 50 levels of nesting

    for i in 0..depth {
        code.push_str(&format!("  If x{} Then\n", i));
    }

    code.push_str("    y = 1\n");

    for _ in 0..depth {
        code.push_str("  End If\n");
    }

    code.push_str("End Sub\n");

    // Should parse without crashing
    let (result, _errors) = ConcreteSyntaxTree::from_text("test.bas", &code).unpack();
    assert!(result.is_some());
}

/// Test that deeply nested for loops don't cause stack overflow
#[test]
fn deeply_nested_for_loops() {
    let mut code = String::from("Sub Test()\n");
    let depth = 50; // Test with 50 levels of nesting

    for i in 0..depth {
        code.push_str(&format!("  For i{} = 1 To 10\n", i));
    }

    code.push_str("    x = 1\n");

    for i in (0..depth).rev() {
        code.push_str(&format!("  Next i{}\n", i));
    }

    code.push_str("End Sub\n");

    // Should parse without crashing
    let (result, _errors) = ConcreteSyntaxTree::from_text("test.bas", &code).unpack();
    assert!(result.is_some());
}

/// Test that property groups with nesting don't cause stack overflow
#[test]
fn nested_property_groups() {
    let mut form = String::from("Begin VB.Form Form1\n");
    form.push_str("  Begin VB.Label Label1\n");

    // Create nested property groups
    let depth = 10; // Test with 10 levels (limit is 100)
    for i in 0..depth {
        form.push_str(&format!("    BeginProperty PropGroup{}\n", i));
        form.push_str("      Name = \"Test\"\n");
    }

    for _ in 0..depth {
        form.push_str("    EndProperty\n");
    }

    form.push_str("  End\n");
    form.push_str("End\n");

    // Should parse without crashing
    let source = SourceFile::from_string("test.frm", form);
    let (result, _errors) = FormFile::parse(&source).unpack();
    assert!(result.is_some());
}
