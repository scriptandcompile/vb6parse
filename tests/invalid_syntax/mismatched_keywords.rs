use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test Sub with End Function mismatch
#[test]
fn sub_with_end_function() {
    let source = r"
Sub TestSub()
    Dim x As Integer
    x = 10
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for sub_with_end_function ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("sub_with_end_function_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("sub_with_end_function_failures", failure_messages);
}

/// Test Function with End Sub mismatch
#[test]
fn function_with_end_sub() {
    let source = r"
Function Calculate(x As Integer) As Integer
    Calculate = x * 2
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for function_with_end_sub ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("function_with_end_sub_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("function_with_end_sub_failures", failure_messages);
}

/// Test Property Get with End Sub mismatch
#[test]
fn property_get_with_end_sub() {
    let source = r"
Property Get Name() As String
    Name = m_name
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for property_get_with_end_sub ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("property_get_with_end_sub_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("property_get_with_end_sub_failures", failure_messages);
}

/// Test Property Let with End Function mismatch
#[test]
fn property_let_with_end_function() {
    let source = r"
Property Let Value(newValue As Integer)
    m_value = newValue
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for property_let_with_end_function ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("property_let_with_end_function_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("property_let_with_end_function_failures", failure_messages);
}

/// Test If with End Select mismatch
#[test]
fn if_with_end_select() {
    let source = r#"
Sub Test()
    If x > 0 Then
        Debug.Print "Positive"
    End Select
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for if_with_end_select ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("if_with_end_select_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("if_with_end_select_failures", failure_messages);
}

/// Test Select Case with End If mismatch
#[test]
fn select_case_with_end_if() {
    let source = r#"
Sub Test()
    Select Case x
        Case 1
            Debug.Print "One"
        Case 2
            Debug.Print "Two"
    End If
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for select_case_with_end_if ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("select_case_with_end_if_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("select_case_with_end_if_failures", failure_messages);
}

/// Test For with Wend mismatch
#[test]
fn for_with_wend() {
    let source = r"
Sub Test()
    For i = 1 To 10
        Debug.Print i
    Wend
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for for_with_wend ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("for_with_wend_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("for_with_wend_failures", failure_messages);
}

/// Test Do While with Next mismatch
#[test]
fn do_while_with_next() {
    let source = r"
Sub Test()
    Do While x > 0
        x = x - 1
    Next
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for do_while_with_next ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("do_while_with_next_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("do_while_with_next_failures", failure_messages);
}

/// Test While with Loop mismatch
#[test]
fn while_with_loop() {
    let source = r"
Sub Test()
    While x > 0
        x = x - 1
    Loop
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for while_with_loop ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("while_with_loop_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("while_with_loop_failures", failure_messages);
}

/// Test Type with End Enum mismatch
#[test]
fn type_with_end_enum() {
    let source = r"
Type Point
    X As Long
    Y As Long
End Enum
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for type_with_end_enum ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/mismatched_keywords");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("type_with_end_enum_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("type_with_end_enum_failures", failure_messages);
}
