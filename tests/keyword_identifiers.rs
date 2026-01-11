//! Tests for keywords used as identifiers in various positions
//!
//! VB6 allows keywords to be used as identifiers (variable names, procedure names, etc.)
//! in most contexts. This test file verifies that keywords are properly converted to
//! Identifier tokens when they appear in identifier positions.

use vb6parse::*;

#[test]
fn keyword_as_sub_name() {
    let source = "Sub Text()\nEnd Sub\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_as_function_name() {
    let source = "Function Database() As String\nEnd Function\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_as_property_name() {
    let source = "Property Get Binary() As Integer\nEnd Property\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_as_variable_in_assignment() {
    let source = "text = \"hello\"\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_as_property_in_assignment() {
    let source = "obj.text = \"hello\"\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn multiple_keywords_as_identifiers() {
    let source = r#"
database = "mydb.mdb"
text = "hello"
obj.binary = True
"#;
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_as_enum_name() {
    let source = "Enum Random\n    Value1\n    Value2\nEnd Enum\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}

#[test]
fn keyword_after_keyword_converted() {
    // Even when a keyword follows another keyword in procedure definition
    let source = "Sub Output()\nEnd Sub\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/keyword_identifiers");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}
