use vb6parse::*;

#[test]
fn conservation_of_whitespace_in_expression() {
    let input = "a = 1 + 2";

    // Use the high-level API to parse directly from source
    let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", input).unpack();
    let cst = cst_opt.expect("CST should have parsed.");
    let tree = cst.to_serializable();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/expression_whitespace");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(tree);
}
