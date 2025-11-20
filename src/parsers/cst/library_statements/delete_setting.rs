use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
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
mod test {
    use crate::*;

    #[test]
    fn deletesetting_with_section_only() {
        // Test DeleteSetting with appname and section (deletes entire section)
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_key() {
        // Test DeleteSetting with appname, section, and key
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup", "Left"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_app_productname() {
        // Test DeleteSetting using App.ProductName
        let source = r#"
Sub Test()
    DeleteSetting App.ProductName, "FileFilter"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_constants() {
        // Test DeleteSetting with constants
        let source = r#"
Sub Test()
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.matches("DeleteSettingStatement").count() >= 3);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_concatenation() {
        // Test DeleteSetting with string concatenation
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Section" & Num, "Key" & Index
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_function_call() {
        // Test DeleteSetting with function call as argument
        let source = r#"
Sub Test()
    DeleteSetting GetAppName(), GetSection(), GetKey()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_parentheses() {
        // Test DeleteSetting with parentheses around arguments
        let source = r#"
Sub Test()
    DeleteSetting ("MyApp"), ("Settings"), ("WindowState")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }
}
