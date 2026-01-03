use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 DeleteSetting statement syntax:
    // - DeleteSetting appname, section[, key]
    //
    // Deletes a section or key setting from an application's entry in the Windows registry.
    //
    // The DeleteSetting statement syntax has these named arguments:
    //
    // | Part     | Description |
    // |----------|-------------|
    // | appname  | Required. String expression containing the name of the application or project to which the section or key setting applies. |
    // | section  | Required. String expression containing the name of the section from which the key setting is being deleted. If only appname and section are provided, the specified section is deleted along with all related key settings. |
    // | key      | Optional. String expression containing the name of the key setting being deleted. |
    //
    // Examples:
    // - DeleteSetting "MyApp", "Startup" (deletes entire Startup section)
    // - DeleteSetting "MyApp", "Startup", "Left" (deletes Left key from Startup section)
    // - DeleteSetting App.ProductName, "FileFilter" (deletes FileFilter section)
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletesetting-statement)
    pub(crate) fn parse_delete_setting_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::DeleteSettingStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn deletesetting_with_section_only() {
        // Test DeleteSetting with appname and section (deletes entire section)
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Startup\""),
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
    fn deletesetting_with_key() {
        // Test DeleteSetting with appname, section, and key
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup", "Left"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Startup\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Left\""),
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
    fn deletesetting_with_app_productname() {
        // Test DeleteSetting using App.ProductName
        let source = r#"
Sub Test()
    DeleteSetting App.ProductName, "FileFilter"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("ProductName"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"FileFilter\""),
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
    fn deletesetting_with_constants() {
        // Test DeleteSetting with constants
        let source = r#"
Sub Test()
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("REGISTRY_KEY"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Settings\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"frmPost.Left\""),
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
    fn deletesetting_multiple_calls() {
        // Test multiple DeleteSetting calls
        let source = r#"
Sub Test()
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Top"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Height"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("REGISTRY_KEY"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Settings\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"frmPost.Left\""),
                        Newline,
                    },
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("REGISTRY_KEY"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Settings\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"frmPost.Top\""),
                        Newline,
                    },
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("REGISTRY_KEY"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Settings\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"frmPost.Height\""),
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
    fn deletesetting_with_variables() {
        // Test DeleteSetting with variables
        let source = r#"
Sub Test()
    Dim appName As String
    Dim sectionName As String
    appName = "MyApp"
    sectionName = "Settings"
    DeleteSetting appName, sectionName
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("appName"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("sectionName"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("appName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"MyApp\""),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sectionName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"Settings\""),
                        },
                        Newline,
                    },
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("appName"),
                        Comma,
                        Whitespace,
                        Identifier ("sectionName"),
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
    fn deletesetting_in_loop() {
        // Test DeleteSetting in a loop
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        DeleteSetting "MyApp", "Item" & i
    Next i
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
                            DeleteSettingStatement {
                                Whitespace,
                                DeleteSettingKeyword,
                                Whitespace,
                                StringLiteral ("\"MyApp\""),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"Item\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("i"),
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
    fn deletesetting_with_concatenation() {
        // Test DeleteSetting with string concatenation
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Section" & Num, "Key" & Index
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Section\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Num"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Key\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Index"),
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
    fn deletesetting_in_if_statement() {
        // Test DeleteSetting in conditional
        let source = r#"
Sub Test()
    If ResetSettings Then
        DeleteSetting "MyApp", "Preferences"
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("ResetSettings"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            DeleteSettingStatement {
                                Whitespace,
                                DeleteSettingKeyword,
                                Whitespace,
                                StringLiteral ("\"MyApp\""),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"Preferences\""),
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
    fn deletesetting_with_function_call() {
        // Test DeleteSetting with function call as argument
        let source = r"
Sub Test()
    DeleteSetting GetAppName(), GetSection(), GetKey()
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        Identifier ("GetAppName"),
                        LeftParenthesis,
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("GetSection"),
                        LeftParenthesis,
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("GetKey"),
                        LeftParenthesis,
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
    fn deletesetting_with_parentheses() {
        // Test DeleteSetting with parentheses around arguments
        let source = r#"
Sub Test()
    DeleteSetting ("MyApp"), ("Settings"), ("WindowState")
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        LeftParenthesis,
                        StringLiteral ("\"MyApp\""),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        LeftParenthesis,
                        StringLiteral ("\"Settings\""),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        LeftParenthesis,
                        StringLiteral ("\"WindowState\""),
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
    fn deletesetting_with_error_handling() {
        // Test DeleteSetting with error handling
        let source = r#"
Sub Test()
    On Error Resume Next
    DeleteSetting "MyApp", "Settings"
    If Err Then MsgBox "Error deleting setting"
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
                    DeleteSettingStatement {
                        Whitespace,
                        DeleteSettingKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Settings\""),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Err"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error deleting setting\""),
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
