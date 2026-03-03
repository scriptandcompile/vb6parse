use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test missing comma between parameters
#[test]
fn missing_comma_between_parameters() {
    let source = r"
Sub Test(x As Integer y As String)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_comma_between_parameters ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_comma_between_parameters_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("missing_comma_between_parameters_failures", failure_messages);
}

/// Test trailing comma in parameter list
#[test]
fn trailing_comma_in_parameters() {
    let source = r"
Function Calculate(x As Integer, y As String,) As Integer
    Calculate = 0
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for trailing_comma_in_parameters ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("trailing_comma_in_parameters_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("trailing_comma_in_parameters_failures", failure_messages);
}

/// Test missing parameter after comma
#[test]
fn missing_parameter_after_comma() {
    let source = r"
Sub Test(x As Integer, )
    Dim result As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for missing_parameter_after_comma ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("missing_parameter_after_comma_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("missing_parameter_after_comma_failures", failure_messages);
}

/// Test duplicate ByVal modifier
#[test]
fn duplicate_byval_modifier() {
    let source = r"
Sub Test(ByVal ByVal x As Integer)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for duplicate_byval_modifier ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("duplicate_byval_modifier_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("duplicate_byval_modifier_failures", failure_messages);
}

/// Test conflicting ByVal and ByRef modifiers
#[test]
fn conflicting_byval_byref() {
    let source = r"
Function Calculate(ByVal ByRef x As Integer) As Integer
    Calculate = x * 2
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for conflicting_byval_byref ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("conflicting_byval_byref_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("conflicting_byval_byref_failures", failure_messages);
}

/// Test Optional parameter before required parameter
#[test]
fn optional_before_required() {
    let source = r"
Sub Test(Optional x As Integer, y As String)
    Debug.Print y
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for optional_before_required ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("optional_before_required_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("optional_before_required_failures", failure_messages);
}

/// Test ParamArray not as last parameter
#[test]
fn paramarray_not_last() {
    let source = r"
Sub Test(ParamArray args() As Variant, x As Integer)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for paramarray_not_last ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("paramarray_not_last_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("paramarray_not_last_failures", failure_messages);
}

/// Test ParamArray with ByVal modifier (not allowed)
#[test]
fn paramarray_with_byval() {
    let source = r"
Sub Test(ByVal ParamArray args() As Variant)
    Dim i As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for paramarray_with_byval ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("paramarray_with_byval_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("paramarray_with_byval_failures", failure_messages);
}

/// Test multiple consecutive commas
#[test]
fn multiple_consecutive_commas() {
    let source = r"
Function Calculate(x As Integer,, y As String) As Integer
    Calculate = 0
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for multiple_consecutive_commas ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("multiple_consecutive_commas_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("multiple_consecutive_commas_failures", failure_messages);
}

/// Test parameter with missing As keyword
#[test]
fn parameter_missing_as_keyword() {
    let source = r"
Sub Test(x Integer, y As String)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for parameter_missing_as_keyword ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("parameter_missing_as_keyword_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("parameter_missing_as_keyword_failures", failure_messages);
}

/// Test Optional with both ByVal and default value
#[test]
fn optional_byval_with_default() {
    let source = r"
Sub Test(Optional ByVal x As Integer = )
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for optional_byval_with_default ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("optional_byval_with_default_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("optional_byval_with_default_failures", failure_messages);
}

/// Test duplicate Optional modifier
#[test]
fn duplicate_optional_modifier() {
    let source = r"
Function Calculate(Optional Optional x As Integer) As Integer
    Calculate = x
End Function
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for duplicate_optional_modifier ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("duplicate_optional_modifier_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("duplicate_optional_modifier_failures", failure_messages);
}

/// Test ParamArray without array parentheses
#[test]
fn paramarray_without_parentheses() {
    let source = r"
Sub Test(ParamArray args As Variant)
    Dim i As Integer
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for paramarray_without_parentheses ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("paramarray_without_parentheses_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("paramarray_without_parentheses_failures", failure_messages);
}

/// Test parameter with type character instead of As clause
#[test]
fn parameter_type_character_with_as() {
    let source = r"
Sub Test(x% As Integer)
    Debug.Print x
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for parameter_type_character_with_as ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_parameter_list");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("parameter_type_character_with_as_cst", tree);

    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("parameter_type_character_with_as_failures", failure_messages);
}
