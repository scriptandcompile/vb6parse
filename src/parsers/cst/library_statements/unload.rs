//! # Unload Statement
//!
//! Removes a form or control from memory.
//!
//! ## Syntax
//!
//! ```vb
//! Unload object
//! ```
//!
//! ## Parts
//!
//! - **object**: Required. An object expression that evaluates to a Form or control. If object is
//!   a form, unloading the form causes all controls on the form to be unloaded as well.
//!
//! ## Remarks
//!
//! - **Form Unloading**: When a form is unloaded, all of its controls are removed from memory and
//!   all values of the form's properties are lost. You can use the `Hide` method to make a form
//!   invisible without unloading it, allowing you to continue to access properties of the form
//!   and its controls.
//! - **Control Arrays**: When you unload a control created at run time with the `Load` statement,
//!   the control is removed from the control array, and the array's upper bound is decremented by one.
//! - **QueryUnload Event**: Before a form is unloaded, the `QueryUnload` event procedure is called.
//!   Setting the `Cancel` argument to `True` in the `QueryUnload` event prevents the form from
//!   being unloaded.
//! - **Unload Event**: After the `QueryUnload` event, the `Unload` event procedure is called. You
//!   can include code in this event procedure to save data or clean up resources.
//! - **Me Keyword**: Within a form's code, you can use `Unload Me` to unload the form itself.
//! - **Subsequent References**: Any subsequent references to properties or controls on an unloaded
//!   form will cause the form to be reloaded and its `Load` event to fire.
//!
//! ## Examples
//!
//! ### Simple Form Unload
//!
//! ```vb
//! Unload Form1
//! ```
//!
//! ### Unload Current Form
//!
//! ```vb
//! Private Sub cmdClose_Click()
//!     Unload Me
//! End Sub
//! ```
//!
//! ### Unload Control Array Element
//!
//! ```vb
//! Unload txtDynamic(5)
//! ```
//!
//! ### Unload With Cleanup
//!
//! ```vb
//! Private Sub Form_Unload(Cancel As Integer)
//!     ' Save data before closing
//!     SaveSettings
//!     CloseDatabase
//! End Sub
//! ```
//!
//! ### Conditional Unload
//!
//! ```vb
//! If UserConfirmed Then
//!     Unload frmDialog
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Save Data Before Unload
//!
//! ```vb
//! Private Sub Form_Unload(Cancel As Integer)
//!     If DataModified Then
//!         Dim response As VbMsgBoxResult
//!         response = MsgBox("Save changes?", vbYesNoCancel)
//!         If response = vbYes Then
//!             SaveData
//!         ElseIf response = vbCancel Then
//!             Cancel = True ' Prevent unload
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ### Unload Multiple Forms
//!
//! ```vb
//! Sub CloseAllForms()
//!     Dim frm As Form
//!     For Each frm In Forms
//!         If frm.Name <> "frmMain" Then
//!             Unload frm
//!         End If
//!     Next frm
//! End Sub
//! ```
//!
//! ### Unload Dynamically Created Controls
//!
//! ```vb
//! Dim i As Integer
//! For i = 1 To 10
//!     Unload lblDynamic(i)
//! Next i
//! ```
//!
//! ## Best Practices
//!
//! 1. **Use Unload vs Hide**: Use `Unload` when you're done with a form and want to free memory.
//!    Use `Hide` when you want to make a form invisible but may need to show it again soon.
//! 2. **Clean Up Resources**: Use the `Unload` event to close database connections, release objects,
//!    and perform other cleanup tasks.
//! 3. **Prevent Accidental Closes**: Use the `QueryUnload` event with `Cancel = True` to prevent
//!    forms from being unloaded when necessary.
//! 4. **Main Form Considerations**: Unloading the startup form (main form) terminates the application
//!    unless you've specified a Sub Main procedure.
//! 5. **Memory Management**: Unloading forms and controls frees memory, which is important in
//!    applications that create many forms or controls dynamically.
//!
//! ## Important Notes
//!
//! - Unloading a form removes it from memory completely
//! - Any data stored in form-level variables is lost
//! - Controls on an unloaded form cannot be accessed
//! - The `Unload` event fires before the form is actually removed
//! - MDI child forms are unloaded when the MDI parent is unloaded
//! - You cannot unload a control that wasn't created with the `Load` statement
//!
//! ## See Also
//!
//! - `Load` statement (loads a form or control into memory)
//! - `Show` method (displays a form)
//! - `Hide` method (hides a form without unloading)
//! - `QueryUnload` event (fires before a form is unloaded)
//! - `Unload` event (fires when a form is being unloaded)
//!
//! ## References
//!
//! - [Microsoft Docs: Unload Statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement)

use crate::parsers::SyntaxKind;

use super::super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Visual Basic 6 Unload statement.
    ///
    /// Unload statement syntax:
    /// - Unload object
    ///
    /// Removes a form or control from memory.
    ///
    /// Examples:
    /// ```vb
    /// Unload Form1
    /// Unload Me
    /// Unload frmDialog
    /// Unload txtControl(5)
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement)
    pub(crate) fn parse_unload_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::UnloadStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn unload_simple() {
        let source = r#"
Sub Test()
    Unload Form1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("UnloadKeyword"));
        assert!(debug.contains("Form1"));
    }

    #[test]
    fn unload_at_module_level() {
        let source = "Unload frmMain\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_me() {
        let source = r#"
Private Sub cmdClose_Click()
    Unload Me
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("Me"));
    }

    #[test]
    fn unload_form() {
        let source = r#"
Sub Test()
    Unload frmDialog
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("frmDialog"));
    }

    #[test]
    fn unload_control_array_element() {
        let source = r#"
Sub Test()
    Unload txtControl(5)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("txtControl"));
    }

    #[test]
    fn unload_preserves_whitespace() {
        let source = "    Unload    Form1    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Unload    Form1    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_with_comment() {
        let source = r#"
Sub Test()
    Unload frmAbout ' Close about dialog
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn unload_in_if_statement() {
        let source = r#"
Sub Test()
    If needsClose Then
        Unload frmSettings
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_inline_if() {
        let source = r#"
Sub Test()
    If closeDialog Then Unload frmInput
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn multiple_unload_statements() {
        let source = r#"
Sub Test()
    Unload Form1
    Unload Form2
    Unload Form3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("UnloadStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn unload_then_end() {
        let source = r#"
Sub Test()
    Unload frmSplash
    End
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("frmSplash"));
    }

    #[test]
    fn unload_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Unload frmCustom
    If Err.Number <> 0 Then
        MsgBox "Error unloading form"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_dynamic_control() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Unload lblLabel(i)
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("lblLabel"));
    }

    #[test]
    fn unload_in_select_case() {
        let source = r#"
Sub Test()
    Select Case formType
        Case 1
            Unload frmA
        Case 2
            Unload frmB
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_with_object_reference() {
        let source = r#"
Sub Test()
    Dim frm As Form1
    Set frm = New Form1
    Unload frm
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_mdi_child() {
        let source = r#"
Sub Test()
    Unload frmChild
    frmMain.Arrange vbCascade
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_in_loop() {
        let source = r#"
Sub Test()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_conditional() {
        let source = r#"
Sub Test()
    If UserConfirmed Then
        Unload frmWarning
    Else
        frmWarning.Show
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_after_hide() {
        let source = r#"
Sub Test()
    frmSettings.Hide
    SaveSettings
    Unload frmSettings
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_with_doevents() {
        let source = r#"
Sub Test()
    Unload frmProgress
    DoEvents
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_array_index_expression() {
        let source = r#"
Sub Test()
    Dim idx As Integer
    idx = 5
    Unload picImage(idx * 2)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("picImage"));
    }

    #[test]
    fn unload_qualified_name() {
        let source = r#"
Sub Test()
    Unload MyProject.frmCustom
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_button_click() {
        let source = r#"
Private Sub Button_Cancel_Click()
    Unload Me
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("Me"));
    }

    #[test]
    fn unload_form_queryunload() {
        let source = r#"
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmChild
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_all_forms_except_main() {
        let source = r#"
Sub CloseAll()
    Dim f As Form
    For Each f In Forms
        If f.Name <> "frmMain" Then
            Unload f
        End If
    Next f
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_with_cleanup() {
        let source = r#"
Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
    CloseDatabase
    Unload frmHelper
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_splash_screen() {
        let source = r#"
Sub Main()
    frmSplash.Show
    DoEvents
    InitializeApp
    Unload frmSplash
    frmMain.Show
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("frmSplash"));
    }

    #[test]
    fn unload_multiple_instances() {
        let source = r#"
Sub Test()
    Unload Form1(0)
    Unload Form1(1)
    Unload Form1(2)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("UnloadStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn unload_before_end() {
        let source = r#"
Private Sub cmdExit_Click()
    Unload Me
    End
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
    }

    #[test]
    fn unload_with_parentheses() {
        let source = r#"
Sub Test()
    Unload (frmDialog)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnloadStatement"));
        assert!(debug.contains("frmDialog"));
    }
}
