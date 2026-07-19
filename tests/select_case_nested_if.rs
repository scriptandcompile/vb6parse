use vb6parse::*;

#[test]
fn select_case_nested_pre_directive_if() {
    let source = r#"
Sub Test()
    Select Case x
        Case 1
            If x > 0 Then
#If DEBUG Then
                a = 1
#Else
                a = 2
#End If
            End If
        Case Else
            b = 2
    End Select
End Sub
"#;

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    assert!(
        failures.is_empty(),
        "unexpected parse failures: {failures:?}"
    );

    let cst = cst_opt.expect("CST should be parsed");
    let tree = cst.to_serializable();
    let text = format!("{tree:#?}");

    assert!(
        !text.contains("Unknown"),
        "should not contain Unknown tokens, found:\n{text}"
    );
    assert!(
        text.contains("CaseElseClause"),
        "should contain a CaseElseClause node, found:\n{text}"
    );
}
