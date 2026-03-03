use vb6parse::parsers::cst::ConcreteSyntaxTree;

/// Test unclosed string literal
#[test]
fn unclosed_string() {
    let source = r#"
Sub Test()
    Dim message As String
    message = "Hello World
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for unclosed_string ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("unclosed_string_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("unclosed_string_failures", failure_messages);
}

/// Test string with incomplete quote escape
#[test]
fn incomplete_quote_escape() {
    let source = r#"
Sub Test()
    Dim text As String
    text = "She said ""Hello
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for incomplete_quote_escape ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("incomplete_quote_escape_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("incomplete_quote_escape_failures", failure_messages);
}

/// Test invalid numeric literal with multiple decimal points
#[test]
fn multiple_decimal_points() {
    let source = r"
Sub Test()
    Dim value As Double
    value = 123.456.789
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for multiple_decimal_points ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("multiple_decimal_points_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("multiple_decimal_points_failures", failure_messages);
}

/// Test invalid hexadecimal literal
#[test]
fn invalid_hex_literal() {
    let source = r"
Sub Test()
    Dim color As Long
    color = &HGGGG
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for invalid_hex_literal ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("invalid_hex_literal_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("invalid_hex_literal_failures", failure_messages);
}

/// Test invalid octal literal
#[test]
fn invalid_octal_literal() {
    let source = r"
Sub Test()
    Dim perms As Long
    perms = &O999
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for invalid_octal_literal ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("invalid_octal_literal_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("invalid_octal_literal_failures", failure_messages);
}

/// Test invalid date literal - bad month
#[test]
fn invalid_date_month() {
    let source = r"
Sub Test()
    Dim dt As Date
    dt = #13/25/2000#
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for invalid_date_month ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("invalid_date_month_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("invalid_date_month_failures", failure_messages);
}

/// Test unclosed date literal
#[test]
fn unclosed_date_literal() {
    let source = r"
Sub Test()
    Dim dt As Date
    dt = #12/25/2000
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for unclosed_date_literal ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("unclosed_date_literal_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("unclosed_date_literal_failures", failure_messages);
}

/// Test invalid exponent in scientific notation
#[test]
fn invalid_scientific_notation() {
    let source = r"
Sub Test()
    Dim value As Double
    value = 1.23E
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for invalid_scientific_notation ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("invalid_scientific_notation_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("invalid_scientific_notation_failures", failure_messages);
}

/// Test number with invalid type suffix
#[test]
fn invalid_number_suffix() {
    let source = r"
Sub Test()
    Dim value As Integer
    value = 123Q
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for invalid_number_suffix ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("invalid_number_suffix_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("invalid_number_suffix_failures", failure_messages);
}

/// Test number with leading zeros (potentially ambiguous)
#[test]
fn number_with_leading_zeros() {
    let source = r"
Sub Test()
    Dim value As Integer
    value = 0123
End Sub
";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    eprintln!("=== Failures for number_with_leading_zeros ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/invalid_syntax/invalid_literals");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();

    insta::assert_yaml_snapshot!("number_with_leading_zeros_cst", tree);

    let failure_messages: Vec<String> = failures.iter().map(|f| format!("{:?}", f)).collect();
    insta::assert_yaml_snapshot!("number_with_leading_zeros_failures", failure_messages);
}
