use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

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
mod tests {
    use crate::assert_tree;
    use crate::*; // Load statement tests

    #[test]
    fn load_simple() {
        let source = r"
Sub Test()
    Load Form1
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
    fn load_at_module_level() {
        let source = "Load frmMain\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            LoadStatement {
                LoadKeyword,
                Whitespace,
                Identifier ("frmMain"),
                Newline,
            },
        ]);
    }

    #[test]
    fn load_form() {
        let source = r"
Sub Test()
    Load frmDialog
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
    fn load_control_array_element() {
        let source = r"
Sub Test()
    Load txtControl(5)
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
    fn load_preserves_whitespace() {
        let source = "    Load    Form1    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            LoadStatement {
                LoadKeyword,
                Whitespace,
                Identifier ("Form1"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn load_with_comment() {
        let source = r"
Sub Test()
    Load frmAbout ' Show about dialog
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
    fn load_in_if_statement() {
        let source = r"
Sub Test()
    If needsForm Then
        Load frmSettings
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
                            Identifier ("needsForm"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            LoadStatement {
                                Whitespace,
                                LoadKeyword,
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
    fn load_inline_if() {
        let source = r"
Sub Test()
    If showDialog Then Load frmInput
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
                            Identifier ("showDialog"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        LoadStatement {
                            LoadKeyword,
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
    fn multiple_load_statements() {
        let source = r"
Sub Test()
    Load Form1
    Load Form2
    Load Form3
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
                        Whitespace,
                        Identifier ("Form1"),
                        Newline,
                    },
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
                        Whitespace,
                        Identifier ("Form2"),
                        Newline,
                    },
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
    fn load_then_show() {
        let source = r"
Sub Test()
    Load frmSplash
    frmSplash.Show
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
                        Whitespace,
                        Identifier ("frmSplash"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("frmSplash"),
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
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
                                StringLiteral ("\"Error loading form\""),
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
    fn load_dynamic_control() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Load lblLabel(i)
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
                            LoadStatement {
                                Whitespace,
                                LoadKeyword,
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
    fn load_before_accessing_controls() {
        let source = r#"
Sub Test()
    Load frmData
    frmData.txtName.Text = "John"
    frmData.Show
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
                    LoadStatement {
                        Whitespace,
                        LoadKeyword,
                        Whitespace,
                        Identifier ("frmData"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            MemberAccessExpression {
                                Identifier ("frmData"),
                                PeriodOperator,
                                Identifier ("txtName"),
                            },
                            PeriodOperator,
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"John\""),
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("frmData"),
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
}
