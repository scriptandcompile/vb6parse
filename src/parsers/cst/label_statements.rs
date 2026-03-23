//! Tests for VB6 label statements and their interaction with control flow.
//!
//! Labels in VB6 are identifiers followed by a colon, used primarily with `GoTo` and `GoSub`.
//! This test file verifies that labels are properly parsed and that End statements
//! immediately following labels are correctly recognized.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn label_followed_by_end_function() {
        let source = r"
Function TestFunction() As Long
    TestFunction = 1
    Exit Function
ExitEnd:
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let text = format!("{:#?}", tree);
        assert!(
            !text.contains("Unknown"),
            "Should not contain Unknown tokens, but found:\n{}",
            text
        );

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/label_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_followed_by_end_sub() {
        let source = r"
Sub TestSub()
    Exit Sub
CleanupLabel:
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let text = format!("{:#?}", tree);
        assert!(
            !text.contains("Unknown"),
            "Should not contain Unknown tokens"
        );

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/label_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn label_followed_by_end_property() {
        let source = r#"
Property Get TestProperty() As String
    TestProperty = "test"
    Exit Property
ErrorHandler:
End Property
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let text = format!("{:#?}", tree);
        assert!(
            !text.contains("Unknown"),
            "Should not contain Unknown tokens"
        );

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/label_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_labels_in_function() {
        let source = r"
Function MultiLabel() As Integer
    On Error GoTo ErrorHandler
    MultiLabel = 1
    Exit Function
ErrorHandler:
    MultiLabel = -1
    GoTo CleanupLabel
CleanupLabel:
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let text = format!("{:#?}", tree);
        assert!(
            !text.contains("Unknown"),
            "Should not contain Unknown tokens"
        );

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/label_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
