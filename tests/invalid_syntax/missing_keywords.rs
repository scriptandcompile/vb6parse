use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test missing Then keyword in If statement
#[test]
fn missing_then_in_if() {
    let source = r"
Sub Test()
    If x > 0
        Debug.Print x
    End If
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_then_in_if ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_then_in_if_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_then_in_if_failures", failure_messages);
}

/// Test missing To keyword in For loop
#[test]
fn missing_to_in_for() {
    let source = r"
Sub Test()
    For i = 1 10
        Debug.Print i
    Next i
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_to_in_for ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_to_in_for_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_to_in_for_failures", failure_messages);
}

/// Test missing As keyword in Dim statement
#[test]
fn missing_as_in_dim() {
    let source = r"
Sub Test()
    Dim x Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_as_in_dim ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_as_in_dim_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_as_in_dim_failures", failure_messages);
}

/// Test missing = in Const declaration
#[test]
fn missing_equals_in_const() {
    let source = r"
Sub Test()
    Const MAX_VALUE 100
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_equals_in_const ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_equals_in_const_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_equals_in_const_failures", failure_messages);
}

/// Test missing Case keyword in Select statement
#[test]
fn missing_case_in_select() {
    let source = r#"
Sub Test()
    Select Case x
        1
            Debug.Print "One"
        Case 2
            Debug.Print "Two"
    End Select
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_case_in_select ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_case_in_select_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_case_in_select_failures", failure_messages);
}

/// Test missing Loop keyword in Do statement
#[test]
fn missing_loop_in_do() {
    let source = r"
Sub Test()
    Do While x > 0
        x = x - 1
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_loop_in_do ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_loop_in_do_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_loop_in_do_failures", failure_messages);
}

/// Test missing Next keyword in For loop
#[test]
fn missing_next_in_for() {
    let source = r"
Sub Test()
    For i = 1 To 10
        Debug.Print i
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_next_in_for ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/missing_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_next_in_for_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("missing_next_in_for_failures", failure_messages);
}
