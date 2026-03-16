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

/// Test that extremely deeply nested controls are handled.
///
/// With iterative parsing using explicit stacks, this handles
/// arbitrary nesting depths (1100+ levels) without stack overflow.
#[test]
fn extremely_nested_controls_depth_limit() {
    // Create a form with very deep nesting
    let mut form = String::from("Begin VB.Form Form1\n");
    let depth = 1100;

    for i in 0..depth {
        form.push_str(&format!("  Begin VB.Frame Frame{}\n", i));
    }

    for _ in 0..depth {
        form.push_str("  End\n");
    }

    form.push_str("End\n");

    // Should parse successfully with iterative parsing
    let source = SourceFile::from_string("test.frm", form);
    let (result, _errors) = FormFile::parse(&source).unpack();

    // Verify we successfully parsed despite extreme nesting
    assert!(result.is_some());
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

/// Test that extremely deeply nested expressions work with iterative parsing.
///
/// Iterative expression parsing uses explicit stacks, so arbitrary
/// nesting depths (2000+ levels) are supported without stack overflow.
#[test]
fn extremely_nested_expressions_no_limit() {
    let mut code = String::from("Sub Test()\n");
    code.push_str("  x = ");

    // This depth would have caused stack overflow with recursive parsing
    let depth = 2000;
    for _ in 0..depth {
        code.push('(');
    }

    code.push_str("42");

    for _ in 0..depth {
        code.push(')');
    }

    code.push_str("\nEnd Sub\n");

    // Should parse successfully with iterative expression parsing
    let (result, _errors) = ConcreteSyntaxTree::from_text("test.bas", &code).unpack();
    assert!(
        result.is_some(),
        "Expression parsing should handle arbitrary depth with iterative approach"
    );
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
    let depth = 10;
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

/// Test that extremely deeply nested property groups are handled.
#[test]
fn extremely_nested_property_groups() {
    let mut form = String::from("Begin VB.Form Form1\n");
    form.push_str("  Begin VB.Label Label1\n");

    // Create extremely nested property groups
    let depth = 150;
    for i in 0..depth {
        form.push_str(&format!("    BeginProperty PropGroup{}\n", i));
        form.push_str("      Name = \"Test\"\n");
    }

    for _ in 0..depth {
        form.push_str("    EndProperty\n");
    }

    form.push_str("  End\n");
    form.push_str("End\n");

    // Should parse without crashing thanks to Phase 2 iterative parsing
    let source = SourceFile::from_string("test.frm", form);
    let (result, _errors) = FormFile::parse(&source).unpack();
    assert!(result.is_some());
}
