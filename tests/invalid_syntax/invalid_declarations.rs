use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test Dim statement with missing identifier
#[test]
fn dim_missing_identifier() {
    let source = r"
Sub Test()
    Dim As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for dim_missing_identifier ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("dim_missing_identifier_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("dim_missing_identifier_failures", failure_messages);
}

/// Test Dim statement with missing type annotation
#[test]
fn dim_missing_type() {
    let source = r"
Sub Test()
    Dim x As
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for dim_missing_type ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("dim_missing_type_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("dim_missing_type_failures", failure_messages);
}

/// Test Function with missing return type
#[test]
fn function_missing_return_type() {
    let source = r"
Function Calculate(x As Integer) As
    Calculate = x * 2
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for function_missing_return_type ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("function_missing_return_type_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("function_missing_return_type_failures", failure_messages);
}

/// Test Sub with missing identifier (empty parentheses)
#[test]
fn sub_missing_name() {
    let source = r#"
Sub ()
    Debug.Print "Test"
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for sub_missing_name ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("sub_missing_name_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("sub_missing_name_failures", failure_messages);
}

/// Test duplicate Public modifiers
#[test]
fn duplicate_public_modifier() {
    let source = r#"
Public Public Sub Test()
    Debug.Print "Duplicate modifier"
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for duplicate_public_modifier ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("duplicate_public_modifier_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("duplicate_public_modifier_failures", failure_messages);
}

/// Test conflicting visibility modifiers
#[test]
fn conflicting_visibility_modifiers() {
    let source = r#"
Public Private Sub Test()
    Debug.Print "Conflicting modifiers"
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for conflicting_visibility_modifiers ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("conflicting_visibility_modifiers_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!(
        "conflicting_visibility_modifiers_failures",
        failure_messages
    );
}

/// Test array declaration with missing bounds
#[test]
fn array_missing_bounds() {
    let source = r"
Sub Test()
    Dim arr() As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for array_missing_bounds ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("array_missing_bounds_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("array_missing_bounds_failures", failure_messages);
}

/// Test Const without value assignment
#[test]
fn const_missing_value() {
    let source = r"
Sub Test()
    Const MAX_VALUE As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for const_missing_value ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("const_missing_value_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("const_missing_value_failures", failure_messages);
}

/// Test Type with missing member name
#[test]
fn type_missing_member_name() {
    let source = r"
Type Point
    As Long
    Y As Long
End Type
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for type_missing_member_name ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("type_missing_member_name_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("type_missing_member_name_failures", failure_messages);
}

/// Test parameter with missing name
#[test]
fn parameter_missing_name() {
    let source = r"
Sub Test(As Integer, y As String)
    Debug.Print y
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for parameter_missing_name ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("parameter_missing_name_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("parameter_missing_name_failures", failure_messages);
}

/// Test Optional parameter without default value
#[test]
fn optional_parameter_missing_default() {
    let source = r"
Sub Test(Optional x As Integer =)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for optional_parameter_missing_default ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("optional_parameter_missing_default_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!(
        "optional_parameter_missing_default_failures",
        failure_messages
    );
}

/// Test duplicate Static modifiers on variables
#[test]
fn duplicate_static_modifier() {
    let source = r"
Sub Test()
    Static Static x As Integer
    x = 10
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for duplicate_static_modifier ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("duplicate_static_modifier_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("duplicate_static_modifier_failures", failure_messages);
}

/// Test Enum with missing member value after equals
#[test]
fn enum_missing_member_value() {
    let source = r"
Enum Colors
    Red = 1
    Green =
    Blue = 3
End Enum
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for enum_missing_member_value ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_declarations");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("enum_missing_member_value_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("enum_missing_member_value_failures", failure_messages);
}
