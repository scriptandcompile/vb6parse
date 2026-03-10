use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test deeply nested If statements
/// This tests indirect recursion through parse_statement_list() and parse_if_statement()
///
/// Note: This test currently uses a moderate depth (50) to test parser behavior.
/// Once recursion depth limits are implemented (see recursion.md, Strategy 5),
/// this test should be updated to verify proper error handling at the limit.
#[test]
fn deeply_nested_if_statements() {
    const DEPTH: usize = 50;

    let mut source = String::from("Sub Test()\n");

    // Generate deeply nested If statements
    for i in 0..DEPTH {
        source.push_str(&format!("{}If x{} Then\n", "    ".repeat(i + 1), i));
    }

    // Add innermost statement
    source.push_str(&format!("{}y = 1\n", "    ".repeat(DEPTH + 1)));

    // Close all If statements
    for i in (0..DEPTH).rev() {
        source.push_str(&format!("{}End If\n", "    ".repeat(i + 1)));
    }

    source.push_str("End Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for deeply_nested_if_statements (depth={}) ===",
        DEPTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("deeply_nested_if_statements_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("deeply_nested_if_statements_failures", failure_messages);
}

/// Test deeply nested For loops
/// This tests indirect recursion through parse_statement_list() and parse_for_statement()
///
/// Note: Uses moderate depth (30) for testing. Once recursion limits are implemented,
/// this should verify proper error handling at MAX_STATEMENT_DEPTH.
#[test]
fn deeply_nested_for_loops() {
    const DEPTH: usize = 30;

    let mut source = String::from("Sub Test()\n");

    // Generate deeply nested For loops
    for i in 0..DEPTH {
        source.push_str(&format!("{}For i{} = 1 To 10\n", "    ".repeat(i + 1), i));
    }

    // Add innermost statement
    source.push_str(&format!("{}x = x + 1\n", "    ".repeat(DEPTH + 1)));

    // Close all For loops
    for i in (0..DEPTH).rev() {
        source.push_str(&format!("{}Next i{}\n", "    ".repeat(i + 1), i));
    }

    source.push_str("End Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for deeply_nested_for_loops (depth={}) ===",
        DEPTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("deeply_nested_for_loops_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("deeply_nested_for_loops_failures", failure_messages);
}

/// Test deeply nested parenthesized expressions
/// This tests direct recursion in parse_expression_with_binding_power()
///
/// Note: Uses moderate depth (100) for testing. Once recursion limits are implemented
/// (Strategy 5, MAX_EXPRESSION_DEPTH = 500), this should verify proper error handling.
#[test]
fn deeply_nested_parentheses() {
    const DEPTH: usize = 100;

    let mut source = String::from("Sub Test()\n    result = ");

    // Generate deeply nested parentheses
    for _ in 0..DEPTH {
        source.push('(');
    }

    source.push_str("x");

    for _ in 0..DEPTH {
        source.push(')');
    }

    source.push_str("\nEnd Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for deeply_nested_parentheses (depth={}) ===",
        DEPTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("deeply_nested_parentheses_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("deeply_nested_parentheses_failures", failure_messages);
}

/// Test long chain of binary operations
/// This tests expression parsing with many infix operators
///
/// Note: Uses moderate length (200) for testing. Once recursion limits are implemented,
/// this should verify the parser handles long expression chains correctly.
#[test]
fn long_binary_operation_chain() {
    const LENGTH: usize = 200;

    let mut source = String::from("Sub Test()\n    result = ");

    // Generate long chain of additions
    for i in 0..LENGTH {
        if i > 0 {
            source.push_str(" + ");
        }
        source.push_str(&format!("x{}", i));
    }

    source.push_str("\nEnd Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for long_binary_operation_chain (length={}) ===",
        LENGTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("long_binary_operation_chain_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("long_binary_operation_chain_failures", failure_messages);
}

/// Test mixed nested control flow (combination of If, For, While, Do, Select)
/// This tests the mutual recursion between various control flow statements
///
/// Note: Uses moderate depth (25) with mixed constructs. Once recursion limits
/// are implemented, verify proper handling of complex nested control flow.
#[test]
fn mixed_nested_control_flow() {
    let source = r#"
Sub Test()
    If a Then
        For i = 1 To 10
            If b Then
                While c
                    Do
                        If d Then
                            For j = 1 To 5
                                Select Case e
                                    Case 1
                                        If f Then
                                            While g
                                                Do While h
                                                    If i Then
                                                        For k = 1 To 3
                                                            x = x + 1
                                                        Next k
                                                    End If
                                                Loop
                                            Wend
                                        End If
                                    Case 2
                                        x = 2
                                End Select
                            Next j
                        End If
                    Loop Until j
                Wend
            End If
        Next i
    End If
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for mixed_nested_control_flow ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("mixed_nested_control_flow_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("mixed_nested_control_flow_failures", failure_messages);
}

/// Test complex nested boolean expression
/// This tests expression recursion with multiple levels of And/Or operations and parentheses
///
/// Note: Once recursion limits are implemented, this should verify proper handling
/// of complex boolean expressions.
#[test]
fn complex_nested_boolean_expression() {
    let source = r#"
Sub Test()
    result = ((a And b) Or (c And d)) And _
             ((e Or f) And (g Or h)) Or _
             (((i And j) Or (k And l)) And _
              ((m Or n) And (o Or p))) And _
             ((q And r) Or ((s And t) And (u Or v)))
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for complex_nested_boolean_expression ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("complex_nested_boolean_expression_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!(
        "complex_nested_boolean_expression_failures",
        failure_messages
    );
}

/// Test deeply nested Select Case statements
/// This tests another form of control flow recursion
///
/// Note: Uses moderate depth (20) for testing. Once recursion limits are implemented,
/// verify proper error handling for deeply nested Select statements.
#[test]
fn deeply_nested_select_case() {
    const DEPTH: usize = 20;

    let mut source = String::from("Sub Test()\n");

    // Generate deeply nested Select Case statements
    for i in 0..DEPTH {
        source.push_str(&format!("{}Select Case x{}\n", "    ".repeat(i + 1), i));
        source.push_str(&format!("{}Case 1\n", "    ".repeat(i + 2)));
    }

    // Add innermost statement
    source.push_str(&format!("{}y = 1\n", "    ".repeat(DEPTH + 2)));

    // Close all Select Case statements
    for i in (0..DEPTH).rev() {
        source.push_str(&format!("{}End Select\n", "    ".repeat(i + 1)));
    }

    source.push_str("End Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for deeply_nested_select_case (depth={}) ===",
        DEPTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("deeply_nested_select_case_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("deeply_nested_select_case_failures", failure_messages);
}

/// Test deeply nested With blocks
/// This tests another recursion pattern in statement list parsing
///
/// Note: Uses moderate depth (25) for testing. Once recursion limits are implemented,
/// verify proper handling of nested With blocks.
#[test]
fn deeply_nested_with_blocks() {
    const DEPTH: usize = 25;

    let mut source = String::from("Sub Test()\n");

    // Generate deeply nested With blocks
    for i in 0..DEPTH {
        source.push_str(&format!("{}With obj{}\n", "    ".repeat(i + 1), i));
    }

    // Add innermost statement
    source.push_str(&format!("{}.Property = 1\n", "    ".repeat(DEPTH + 1)));

    // Close all With blocks
    for i in (0..DEPTH).rev() {
        source.push_str(&format!("{}End With\n", "    ".repeat(i + 1)));
    }

    source.push_str("End Sub\n");

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", &source).unpack();

    eprintln!(
        "=== Failures for deeply_nested_with_blocks (depth={}) ===",
        DEPTH
    );
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("deeply_nested_with_blocks_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("deeply_nested_with_blocks_failures", failure_messages);
}

/// Test combination of nested expressions and statements
/// This tests interaction between expression recursion and statement recursion
///
/// Note: Tests realistic scenario with reasonable nesting. Once recursion limits
/// are implemented, verify both expression and statement depth tracking work correctly.
#[test]
fn combined_expression_and_statement_nesting() {
    let source = r#"
Sub Test()
    If ((a + b) * (c + d)) > ((e - f) / (g - h)) Then
        For i = (x * y) To ((z + w) * 2)
            If (((p And q) Or (r And s)) And ((t Or u) And (v Or w))) Then
                result = ((a + (b * (c + (d * e)))) - ((f * (g + h)) / i))
                While (x > ((y * z) + (w / 2)))
                    Do
                        x = x - ((a + b) * (c + d))
                    Loop Until (x < ((y + z) / 2))
                Wend
            End If
        Next i
    End If
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for combined_expression_and_statement_nesting ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/edge_cases/recursion_limits");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("combined_expression_and_statement_nesting_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!(
        "combined_expression_and_statement_nesting_failures",
        failure_messages
    );
}
