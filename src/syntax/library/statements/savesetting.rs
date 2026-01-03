//! # `SaveSetting` Statement
//!
//! Saves or creates an application entry in the Windows registry or (on the Macintosh) information in the application's initialization file.
//!
//! ## Syntax
//!
//! ```vb
//! SaveSetting appname, section, key, setting
//! ```
//!
//! ## Parts
//!
//! - **appname**: Required. String expression containing the name of the application or project to which the setting applies.
//! - **section**: Required. String expression containing the name of the section in which the key setting is being saved.
//! - **key**: Required. String expression containing the name of the key setting being saved.
//! - **setting**: Required. Expression containing the value to which key is being set.
//!
//! ## Remarks
//!
//! - **Registry Location**: On Windows, `SaveSetting` writes to the registry under the path:
//!   `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section\key`
//! - **String Values**: The setting argument is always stored as a string value in the registry.
//! - **Creating Entries**: If the specified key setting doesn't exist, `SaveSetting` creates it.
//! - **Creating Sections**: If the specified section doesn't exist, `SaveSetting` creates it.
//! - **Application Name**: The appname is typically the name of your application. Multiple applications can use the same registry location by using the same appname.
//! - **Section Organization**: Use sections to organize related settings. For example, you might have a "Startup" section and a "Display" section.
//! - **Type Conversion**: Numeric values and other data types are automatically converted to strings when saved.
//! - **Security**: Settings are stored per user (`HKEY_CURRENT_USER`), not per machine.
//! - **`GetSetting` Function**: Use the `GetSetting` function to retrieve values saved with `SaveSetting`.
//! - **`DeleteSetting` Statement**: Use `DeleteSetting` to remove registry entries created by `SaveSetting`.
//!
//! ## Examples
//!
//! ### Save a Simple Setting
//!
//! ```vb
//! SaveSetting "MyApp", "Startup", "Left", 100
//! SaveSetting "MyApp", "Startup", "Top", 100
//! ```
//!
//! ### Save User Preferences
//!
//! ```vb
//! SaveSetting "MyApp", "Preferences", "BackColor", vbBlue
//! SaveSetting "MyApp", "Preferences", "FontName", "Arial"
//! SaveSetting "MyApp", "Preferences", "FontSize", 12
//! ```
//!
//! ### Save Form Position on Close
//!
//! ```vb
//! Private Sub Form_Unload(Cancel As Integer)
//!     SaveSetting App.Title, "Position", "Left", Me.Left
//!     SaveSetting App.Title, "Position", "Top", Me.Top
//!     SaveSetting App.Title, "Position", "Width", Me.Width
//!     SaveSetting App.Title, "Position", "Height", Me.Height
//! End Sub
//! ```
//!
//! ### Save Boolean Settings
//!
//! ```vb
//! ' Save a boolean as a string
//! SaveSetting "MyApp", "Options", "AutoSave", CStr(chkAutoSave.Value)
//! ```
//!
//! ### Save with Variables
//!
//! ```vb
//! Dim userName As String
//! userName = txtUserName.Text
//! SaveSetting "MyApp", "User", "LastUser", userName
//! ```
//!
//! ### Save Multiple Related Settings
//!
//! ```vb
//! Sub SaveWindowSettings()
//!     Dim appName As String
//!     appName = App.Title
//!     
//!     SaveSetting appName, "Window", "Maximized", Me.WindowState = vbMaximized
//!     SaveSetting appName, "Window", "Visible", Me.Visible
//!     SaveSetting appName, "Window", "Caption", Me.Caption
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Using App.Title for Application Name
//!
//! ```vb
//! ' Ensures consistent application name across all settings
//! SaveSetting App.Title, "Database", "ConnectionString", connStr
//! ```
//!
//! ### Organizing Settings by Feature
//!
//! ```vb
//! ' Group related settings in sections
//! SaveSetting "MyApp", "Display", "Theme", "Dark"
//! SaveSetting "MyApp", "Display", "Language", "English"
//! SaveSetting "MyApp", "Network", "Port", 8080
//! SaveSetting "MyApp", "Network", "Timeout", 30
//! ```
//!
//! ## Important Notes
//!
//! - **Platform Differences**: On Windows, settings are stored in the registry. On other platforms, behavior may vary.
//! - **String Storage**: All values are stored as strings, so you may need to convert them back when retrieving with `GetSetting`.
//! - **Registry Cleanup**: Use `DeleteSetting` to remove settings when they're no longer needed.
//! - **Error Handling**: `SaveSetting` can fail if the registry is locked or permissions are insufficient.
//!
//! ## See Also
//!
//! - `GetSetting` function (retrieve saved settings)
//! - `GetAllSettings` function (retrieve all settings from a section)
//! - `DeleteSetting` statement (delete registry entries)
//!
//! ## References
//!
//! - [SaveSetting Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/savesetting-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `SaveSetting` statement.
    pub(crate) fn parse_savesetting_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::SaveSettingStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn savesetting_simple() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Startup", "Left", 100
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("SaveSettingKeyword"));
    }

    #[test]
    fn savesetting_at_module_level() {
        let source = "SaveSetting \"MyApp\", \"Settings\", \"Value\", \"Data\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_variables() {
        let source = r"
Sub Test()
    SaveSetting appName, sectionName, keyName, value
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("appName"));
        assert!(debug.contains("sectionName"));
        assert!(debug.contains("keyName"));
        assert!(debug.contains("value"));
    }

    #[test]
    fn savesetting_with_app_title() {
        let source = r#"
Sub Test()
    SaveSetting App.Title, "Position", "Left", Me.Left
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("App"));
        assert!(debug.contains("Title"));
    }

    #[test]
    fn savesetting_with_numeric_value() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Display", "Width", 800
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("800"));
    }

    #[test]
    fn savesetting_with_string_value() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "User", "Name", "John Doe"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_control_property() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Settings", "BackColor", Form1.BackColor
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("Form1"));
        assert!(debug.contains("BackColor"));
    }

    #[test]
    fn savesetting_with_concatenation() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "User", "FullName", firstName & " " & lastName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("firstName"));
        assert!(debug.contains("lastName"));
    }

    #[test]
    fn savesetting_with_cstr() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Options", "AutoSave", CStr(chkAutoSave.Value)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("chkAutoSave"));
    }

    #[test]
    fn savesetting_inside_if_statement() {
        let source = r#"
If saveSettings Then
    SaveSetting "MyApp", "Prefs", "Theme", "Dark"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_inside_loop() {
        let source = r#"
For i = 1 To 10
    SaveSetting "MyApp", "Item" & i, "Value", items(i)
Next i
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_comment() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Window", "Left", 100 ' Save window position
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("' Save window position"));
    }

    #[test]
    fn savesetting_preserves_whitespace() {
        let source = "SaveSetting   \"App\"  ,  \"Sec\"  ,  \"Key\"  ,  \"Val\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_form_unload() {
        let source = r#"
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Position", "Left", Me.Left
    SaveSetting App.Title, "Position", "Top", Me.Top
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_multiple_on_same_line() {
        let source =
            "SaveSetting \"A\", \"S\", \"K1\", \"V1\": SaveSetting \"A\", \"S\", \"K2\", \"V2\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_in_select_case() {
        let source = r#"
Select Case mode
    Case 1
        SaveSetting "MyApp", "Mode", "Current", "Simple"
    Case 2
        SaveSetting "MyApp", "Mode", "Current", "Advanced"
End Select
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_format() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "DateTime", "LastRun", Format$(Now, "yyyy-mm-dd")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("Now"));
    }

    #[test]
    fn savesetting_boolean_value() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Options", "Visible", Me.Visible
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("Visible"));
    }

    #[test]
    fn savesetting_in_with_block() {
        let source = r#"
With Form1
    SaveSetting "MyApp", "Form", "Width", .Width
End With
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("Width"));
    }

    #[test]
    fn savesetting_in_sub() {
        let source = r#"
Sub SaveUserPreferences()
    SaveSetting "MyApp", "User", "Name", userName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_in_function() {
        let source = r#"
Function StoreSettings() As Boolean
    SaveSetting "MyApp", "Config", "Version", "1.0"
    StoreSettings = True
End Function
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_array_element() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Colors", "Item" & i, colors(i)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("colors"));
    }

    #[test]
    fn savesetting_in_class_module() {
        let source = r#"
Private appName As String

Public Sub SaveConfig(key As String, value As String)
    SaveSetting appName, "Config", key, value
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_expression() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Math", "Result", x + y * 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_with_textbox() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Input", "UserName", txtUserName.Text
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("txtUserName"));
        assert!(debug.contains("Text"));
    }

    #[test]
    fn savesetting_with_line_continuation() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", _
        "Section", "Key", "Value"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_window_state() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Window", "State", Me.WindowState
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("WindowState"));
    }

    #[test]
    fn savesetting_error_handling() {
        let source = r#"
On Error Resume Next
SaveSetting "MyApp", "Settings", "Value", data
If Err.Number <> 0 Then
    MsgBox "Error saving settings"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
    }

    #[test]
    fn savesetting_nested_sections() {
        let source = r#"
Sub Test()
    SaveSetting "MyApp", "Database\Connection", "Server", serverName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SaveSettingStatement"));
        assert!(debug.contains("serverName"));
    }
}
