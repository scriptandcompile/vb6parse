use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    // VB6 Load statement syntax:
    // - Load object
    //
    // Loads a form or control into memory.
    //
    // The Load statement syntax has this part:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | object        | Required. An object expression that evaluates to a Form or control. |
    //
    // Remarks:
    // - When Visual Basic loads a form, it sets the form's Visible property to False.
    // - After loading a form, you can use the Show method to make the form visible.
    // - The controls on a form aren't accessible until the form is loaded.
    // - Load is typically used with forms that aren't shown at startup or with control arrays.
    // - For control arrays, you must use Load to create controls at run time.
    // - When you load a control array element, Visual Basic automatically increases the array's
    //   upper bound to accommodate the new element.
    // - You can't load a control that doesn't already exist at design time.
    // - The Load event occurs when the form is loaded.
    //
    // Examples:
    // ```vb
    // Load Form1
    // Load frmDialog
    // Load txtControl(5)
    // Load MyForm
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/load-statement)
    pub(crate) fn parse_load_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::LoadStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Load statement tests
    #[test]
    fn load_simple() {
        let source = r"
Sub Test()
    Load Form1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("LoadKeyword"));
        assert!(debug.contains("Form1"));
    }

    #[test]
    fn load_at_module_level() {
        let source = "Load frmMain\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }

    #[test]
    fn load_form() {
        let source = r"
Sub Test()
    Load frmDialog
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("frmDialog"));
    }

    #[test]
    fn load_control_array_element() {
        let source = r"
Sub Test()
    Load txtControl(5)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("txtControl"));
    }

    #[test]
    fn load_preserves_whitespace() {
        let source = "    Load    Form1    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Load    Form1    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }

    #[test]
    fn load_with_comment() {
        let source = r"
Sub Test()
    Load frmAbout ' Show about dialog
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn load_in_if_statement() {
        let source = r"
Sub Test()
    If needsForm Then
        Load frmSettings
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }

    #[test]
    fn load_inline_if() {
        let source = r"
Sub Test()
    If showDialog Then Load frmInput
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }

    #[test]
    fn multiple_load_statements() {
        let source = r"
Sub Test()
    Load Form1
    Load Form2
    Load Form3
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("LoadStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn load_then_show() {
        let source = r"
Sub Test()
    Load frmSplash
    frmSplash.Show
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("frmSplash"));
    }

    #[test]
    fn load_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Load frmCustom
    If Err.Number <> 0 Then
        MsgBox "Error loading form"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }

    #[test]
    fn load_dynamic_control() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Load lblLabel(i)
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
        assert!(debug.contains("lblLabel"));
    }

    #[test]
    fn load_before_accessing_controls() {
        let source = r#"
Sub Test()
    Load frmData
    frmData.txtName.Text = "John"
    frmData.Show
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LoadStatement"));
    }
}
