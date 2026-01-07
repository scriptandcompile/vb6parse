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
//! - **`QueryUnload` Event**: Before a form is unloaded, the `QueryUnload` event procedure is called.
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

use crate::parsers::cst::Parser;

impl Parser<'_> {
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
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn unload_simple() {
        let source = r"
Sub Test()
    Unload Form1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_at_module_level() {
        let source = "Unload frmMain\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            UnloadStatement {
                UnloadKeyword,
                Whitespace,
                Identifier ("frmMain"),
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_me() {
        let source = r"
Private Sub cmdClose_Click()
    Unload Me
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("cmdClose_Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        MeKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_form() {
        let source = r"
Sub Test()
    Unload frmDialog
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmDialog"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_control_array_element() {
        let source = r"
Sub Test()
    Unload txtControl(5)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("txtControl"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_preserves_whitespace() {
        let source = "    Unload    Form1    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            UnloadStatement {
                UnloadKeyword,
                Whitespace,
                Identifier ("Form1"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_with_comment() {
        let source = r"
Sub Test()
    Unload frmAbout ' Close about dialog
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmAbout"),
                        Whitespace,
                        EndOfLineComment,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_in_if_statement() {
        let source = r"
Sub Test()
    If needsClose Then
        Unload frmSettings
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("needsClose"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            UnloadStatement {
                                Whitespace,
                                UnloadKeyword,
                                Whitespace,
                                Identifier ("frmSettings"),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_inline_if() {
        let source = r"
Sub Test()
    If closeDialog Then Unload frmInput
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("closeDialog"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        UnloadStatement {
                            UnloadKeyword,
                            Whitespace,
                            Identifier ("frmInput"),
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn multiple_unload_statements() {
        let source = r"
Sub Test()
    Unload Form1
    Unload Form2
    Unload Form3
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form2"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form3"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_then_end() {
        let source = r"
Sub Test()
    Unload frmSplash
    End
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmSplash"),
                        Newline,
                    },
                    Whitespace,
                    Unknown,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmCustom"),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Error unloading form\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_dynamic_control() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Unload lblLabel(i)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                        StatementList {
                            UnloadStatement {
                                Whitespace,
                                UnloadKeyword,
                                Whitespace,
                                Identifier ("lblLabel"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_in_select_case() {
        let source = r"
Sub Test()
    Select Case formType
        Case 1
            Unload frmA
        Case 2
            Unload frmB
    End Select
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("formType"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                UnloadStatement {
                                    Whitespace,
                                    UnloadKeyword,
                                    Whitespace,
                                    Identifier ("frmA"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                UnloadStatement {
                                    Whitespace,
                                    UnloadKeyword,
                                    Whitespace,
                                    Identifier ("frmB"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_with_object_reference() {
        let source = r"
Sub Test()
    Dim frm As Form1
    Set frm = New Form1
    Unload frm
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("frm"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                    },
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("frm"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frm"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_mdi_child() {
        let source = r"
Sub Test()
    Unload frmChild
    frmMain.Arrange vbCascade
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmChild"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("frmMain"),
                        PeriodOperator,
                        Identifier ("Arrange"),
                        Whitespace,
                        Identifier ("vbCascade"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_in_loop() {
        let source = r"
Sub Test()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("frm"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("Form"),
                        Newline,
                    },
                    ForEachStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        EachKeyword,
                        Whitespace,
                        Identifier ("frm"),
                        Whitespace,
                        InKeyword,
                        Whitespace,
                        Identifier ("Forms"),
                        Newline,
                        StatementList {
                            UnloadStatement {
                                Whitespace,
                                UnloadKeyword,
                                Whitespace,
                                Identifier ("frm"),
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_conditional() {
        let source = r"
Sub Test()
    If UserConfirmed Then
        Unload frmWarning
    Else
        frmWarning.Show
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("UserConfirmed"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            UnloadStatement {
                                Whitespace,
                                UnloadKeyword,
                                Whitespace,
                                Identifier ("frmWarning"),
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("frmWarning"),
                                    PeriodOperator,
                                    Identifier ("Show"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_after_hide() {
        let source = r"
Sub Test()
    frmSettings.Hide
    SaveSettings
    Unload frmSettings
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("frmSettings"),
                        PeriodOperator,
                        Identifier ("Hide"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("SaveSettings"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmSettings"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_with_doevents() {
        let source = r"
Sub Test()
    Unload frmProgress
    DoEvents
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmProgress"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("DoEvents"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_array_index_expression() {
        let source = r"
Sub Test()
    Dim idx As Integer
    idx = 5
    Unload picImage(idx * 2)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("idx"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("idx"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("5"),
                        },
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("picImage"),
                        LeftParenthesis,
                        Identifier ("idx"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("2"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_qualified_name() {
        let source = r"
Sub Test()
    Unload MyProject.frmCustom
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("MyProject"),
                        PeriodOperator,
                        Identifier ("frmCustom"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_button_click() {
        let source = r"
Private Sub Button_Cancel_Click()
    Unload Me
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Button_Cancel_Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        MeKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_form_queryunload() {
        let source = r"
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmChild
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Form_QueryUnload"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("Cancel"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("UnloadMode"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmChild"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("CloseAll"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("f"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("Form"),
                        Newline,
                    },
                    ForEachStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        EachKeyword,
                        Whitespace,
                        Identifier ("f"),
                        Whitespace,
                        InKeyword,
                        Whitespace,
                        Identifier ("Forms"),
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                BinaryExpression {
                                    MemberAccessExpression {
                                        Identifier ("f"),
                                        PeriodOperator,
                                        NameKeyword,
                                    },
                                    Whitespace,
                                    InequalityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"frmMain\""),
                                    },
                                },
                                Whitespace,
                                ThenKeyword,
                                Newline,
                                StatementList {
                                    UnloadStatement {
                                        Whitespace,
                                        UnloadKeyword,
                                        Whitespace,
                                        Identifier ("f"),
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                EndKeyword,
                                Whitespace,
                                IfKeyword,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("f"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_with_cleanup() {
        let source = r"
Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
    CloseDatabase
    Unload frmHelper
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Form_Unload"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("Cancel"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("SaveSettings"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("CloseDatabase"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmHelper"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_splash_screen() {
        let source = r"
Sub Main()
    frmSplash.Show
    DoEvents
    InitializeApp
    Unload frmSplash
    frmMain.Show
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("frmSplash"),
                        PeriodOperator,
                        Identifier ("Show"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("DoEvents"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("InitializeApp"),
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("frmSplash"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("frmMain"),
                        PeriodOperator,
                        Identifier ("Show"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_multiple_instances() {
        let source = r"
Sub Test()
    Unload Form1(0)
    Unload Form1(1)
    Unload Form1(2)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        LeftParenthesis,
                        IntegerLiteral ("0"),
                        RightParenthesis,
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Newline,
                    },
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        LeftParenthesis,
                        IntegerLiteral ("2"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_before_end() {
        let source = r"
Private Sub cmdExit_Click()
    Unload Me
    End
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("cmdExit_Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        MeKeyword,
                        Newline,
                    },
                    Whitespace,
                    Unknown,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn unload_with_parentheses() {
        let source = r"
Sub Test()
    Unload (frmDialog)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    UnloadStatement {
                        Whitespace,
                        UnloadKeyword,
                        Whitespace,
                        LeftParenthesis,
                        Identifier ("frmDialog"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
