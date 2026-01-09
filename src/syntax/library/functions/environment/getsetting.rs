//! # `GetSetting` Function
//!
//! Returns a registry key setting value from the Windows registry.
//!
//! ## Syntax
//!
//! ```vb
//! GetSetting(appname, section, key[, default])
//! ```
//!
//! ## Parameters
//!
//! - `appname` (Required): `String` expression containing the name of the application or project whose key setting is requested. On Windows, this is a subkey under `HKEY_CURRENT_USER\Software\VB and VBA Program Settings`.
//! - `section` (Required): `String` expression containing the name of the section where the key setting is found.
//! - `key` (Required): `String` expression containing the name of the key setting to return.
//! - `default` (Optional): Expression containing the value to return if no value is set in the key setting. If omitted, default is assumed to be a zero-length string ("").
//!
//! ## Return Value
//!
//! Returns a String containing the value of the specified registry key. If the key doesn't exist and no default is provided, returns an empty string.
//!
//! ## Remarks
//!
//! The `GetSetting` function retrieves settings from the Windows registry that were previously saved using the `SaveSetting` statement. The settings are stored in the application's subkey under:
//!
//! `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section`
//!
//! - If the registry key doesn't exist, `GetSetting` returns the default value (or "" if no default specified)
//! - `GetSetting` only works with the `HKEY_CURRENT_USER` registry hive
//! - For more advanced registry access, use Windows API functions like `RegOpenKeyEx` and `RegQueryValueEx`
//! - The `appname`, `section`, and `key` parameters are case-insensitive
//! - `GetSetting` is designed to work with `SaveSetting`, `DeleteSetting`, and `GetAllSettings`
//! - On non-Windows platforms, behavior may vary or be unsupported
//!
//! ## Typical Uses
//!
//! 1. **Application Configuration**: Retrieve user preferences and application settings
//! 2. **User Preferences**: Load window positions, sizes, and UI state
//! 3. **Recent Files**: Get most recently used files or paths
//! 4. **Database Connections**: Retrieve connection strings and server names
//! 5. **Feature Toggles**: Load feature flags and experimental settings
//! 6. **Localization**: Get language and regional preferences
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Get a simple setting with default
//! Dim userName As String
//! userName = GetSetting("MyApp", "User", "Name", "Guest")
//!
//! ' Example 2: Get window position
//! Dim formLeft As String
//! formLeft = GetSetting("MyApp", "Window", "Left", "0")
//!
//! ' Example 3: Get setting without default
//! Dim lastFile As String
//! lastFile = GetSetting("MyApp", "Recent", "File1")
//!
//! ' Example 4: Get database connection
//! Dim connString As String
//! connString = GetSetting("MyApp", "Database", "ConnectionString", "")
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Load form position and size
//! Private Sub Form_Load()
//!     Me.Left = CLng(GetSetting("MyApp", "MainForm", "Left", "0"))
//!     Me.Top = CLng(GetSetting("MyApp", "MainForm", "Top", "0"))
//!     Me.Width = CLng(GetSetting("MyApp", "MainForm", "Width", "6000"))
//!     Me.Height = CLng(GetSetting("MyApp", "MainForm", "Height", "4500"))
//! End Sub
//!
//! ' Pattern 2: Check if setting exists
//! Function SettingExists(app As String, section As String, key As String) As Boolean
//!     Dim marker As String
//!     marker = String$(10, "X")
//!     SettingExists = (GetSetting(app, section, key, marker) <> marker)
//! End Function
//!
//! ' Pattern 3: Get with type conversion
//! Dim showTips As Boolean
//! showTips = CBool(GetSetting("MyApp", "Options", "ShowTips", "True"))
//!
//! ' Pattern 4: Get recent file list
//! Dim i As Integer
//! Dim recentFiles() As String
//! ReDim recentFiles(1 To 10)
//! For i = 1 To 10
//!     recentFiles(i) = GetSetting("MyApp", "Recent", "File" & i, "")
//!     If recentFiles(i) = "" Then Exit For
//! Next i
//!
//! ' Pattern 5: Get connection info
//! Dim server As String, database As String
//! server = GetSetting("MyApp", "Database", "Server", "localhost")
//! database = GetSetting("MyApp", "Database", "Name", "MyDB")
//!
//! ' Pattern 6: Get user preference with validation
//! Dim fontSize As Integer
//! fontSize = CInt(GetSetting("MyApp", "UI", "FontSize", "10"))
//! If fontSize < 8 Or fontSize > 72 Then fontSize = 10
//!
//! ' Pattern 7: Get setting in With block
//! With Form1
//!     .BackColor = CLng(GetSetting("MyApp", "Colors", "Background", "16777215"))
//! End With
//!
//! ' Pattern 8: Conditional loading
//! If GetSetting("MyApp", "Options", "AutoSave", "False") = "True" Then
//!     EnableAutoSave
//! End If
//!
//! ' Pattern 9: Get multiple related settings
//! Dim smtp As String, port As String, useTLS As String
//! smtp = GetSetting("MyApp", "Email", "SMTPServer", "smtp.gmail.com")
//! port = GetSetting("MyApp", "Email", "Port", "587")
//! useTLS = GetSetting("MyApp", "Email", "UseTLS", "True")
//!
//! ' Pattern 10: Safe retrieval with error handling
//! On Error Resume Next
//! Dim value As String
//! value = GetSetting("MyApp", "Config", "Setting", "DefaultValue")
//! If Err.Number <> 0 Then
//!     value = "DefaultValue"
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Settings manager class
//! Public Class AppSettings
//!     Private Const APP_NAME As String = "MyApplication"
//!     
//!     Public Function GetStringSetting(section As String, key As String, _
//!                                      Optional defaultValue As String = "") As String
//!         GetStringSetting = GetSetting(APP_NAME, section, key, defaultValue)
//!     End Function
//!     
//!     Public Function GetIntegerSetting(section As String, key As String, _
//!                                       Optional defaultValue As Integer = 0) As Integer
//!         Dim value As String
//!         value = GetSetting(APP_NAME, section, key, CStr(defaultValue))
//!         On Error Resume Next
//!         GetIntegerSetting = CInt(value)
//!         If Err.Number <> 0 Then GetIntegerSetting = defaultValue
//!     End Function
//!     
//!     Public Function GetBooleanSetting(section As String, key As String, _
//!                                       Optional defaultValue As Boolean = False) As Boolean
//!         Dim value As String
//!         value = GetSetting(APP_NAME, section, key, CStr(defaultValue))
//!         GetBooleanSetting = CBool(value)
//!     End Function
//! End Class
//!
//! ' Example 2: Application configuration loader
//! Private Sub LoadApplicationConfig()
//!     Dim config As New Collection
//!     
//!     config.Add GetSetting("MyApp", "Paths", "Data", App.Path & "\Data"), "DataPath"
//!     config.Add GetSetting("MyApp", "Paths", "Export", App.Path & "\Export"), "ExportPath"
//!     config.Add GetSetting("MyApp", "Paths", "Temp", Environ$("TEMP")), "TempPath"
//!     
//!     config.Add GetSetting("MyApp", "Database", "Server", "localhost"), "DBServer"
//!     config.Add GetSetting("MyApp", "Database", "Name", "AppDB"), "DBName"
//!     
//!     config.Add GetSetting("MyApp", "Options", "AutoBackup", "True"), "AutoBackup"
//!     config.Add GetSetting("MyApp", "Options", "BackupInterval", "60"), "BackupInterval"
//!     
//!     Set g_AppConfig = config
//! End Sub
//!
//! ' Example 3: Multi-user profile system
//! Public Function LoadUserProfile(userName As String) As UserProfile
//!     Dim profile As New UserProfile
//!     Dim section As String
//!     
//!     section = "User_" & userName
//!     
//!     With profile
//!         .FullName = GetSetting("MyApp", section, "FullName", userName)
//!         .Email = GetSetting("MyApp", section, "Email", "")
//!         .Role = GetSetting("MyApp", section, "Role", "User")
//!         .Theme = GetSetting("MyApp", section, "Theme", "Default")
//!         .Language = GetSetting("MyApp", section, "Language", "en-US")
//!         .LastLogin = GetSetting("MyApp", section, "LastLogin", "")
//!     End With
//!     
//!     LoadUserProfile = profile
//! End Function
//!
//! ' Example 4: MRU (Most Recently Used) manager
//! Public Class MRUManager
//!     Private Const MAX_MRU As Integer = 10
//!     Private Const APP_NAME As String = "MyApp"
//!     Private Const SECTION As String = "MRU"
//!     
//!     Public Function GetMRUList() As Collection
//!         Dim mruList As New Collection
//!         Dim i As Integer
//!         Dim item As String
//!         
//!         For i = 1 To MAX_MRU
//!             item = GetSetting(APP_NAME, SECTION, "Item" & i, "")
//!             If Len(item) > 0 Then
//!                 mruList.Add item
//!             Else
//!                 Exit For
//!             End If
//!         Next i
//!         
//!         Set GetMRUList = mruList
//!     End Function
//!     
//!     Public Sub AddMRUItem(filePath As String)
//!         Dim mruList As Collection
//!         Dim i As Integer
//!         Dim item As String
//!         
//!         Set mruList = GetMRUList()
//!         
//!         ' Remove if already exists
//!         For i = 1 To mruList.Count
//!             If StrComp(mruList(i), filePath, vbTextCompare) = 0 Then
//!                 mruList.Remove i
//!                 Exit For
//!             End If
//!         Next i
//!         
//!         ' Add to top
//!         mruList.Add filePath, , 1
//!         
//!         ' Save back
//!         For i = 1 To mruList.Count
//!             If i > MAX_MRU Then Exit For
//!             SaveSetting APP_NAME, SECTION, "Item" & i, mruList(i)
//!         Next i
//!     End Sub
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! `GetSetting` generally doesn't raise errors, but returns the default value if the setting doesn't exist:
//!
//! - **No Error**: If the registry key doesn't exist, returns default (or "" if no default)
//! - **No Error**: If appname, section, or key is empty, returns default value
//! - **Type Mismatch**: Can occur when converting returned String to another type (e.g., `CInt`)
//! - **Registry Access**: On systems where registry access is restricted, may return defaults
//!
//! ```vb
//! ' Safe retrieval with type conversion
//! On Error Resume Next
//! Dim timeout As Integer
//! timeout = CInt(GetSetting("MyApp", "Network", "Timeout", "30"))
//! If Err.Number <> 0 Then
//!     timeout = 30
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - **Registry Access**: Each call accesses the Windows registry, which is slower than memory access
//! - **Caching**: Consider caching frequently used settings in memory
//! - **Startup Time**: Loading many settings at startup can slow application initialization
//! - **Batch Loading**: Use `GetAllSettings` to retrieve all settings in a section at once for better performance
//!
//! ## Best Practices
//!
//! 1. **Use Defaults**: Always provide sensible default values
//! 2. **Validate Values**: Validate retrieved settings before using them
//! 3. **Cache Settings**: Load settings once and cache them for the session
//! 4. **Consistent Naming**: Use consistent naming conventions for `appname`, `section`, and `key`
//! 5. **Error Handling**: Use error handling when converting string values to other types
//! 6. **Cleanup**: Use `DeleteSetting` to remove obsolete settings
//! 7. **Documentation**: Document all registry keys used by your application
//!
//! ## Comparison with Other Registry Functions
//!
//! | Function | Purpose | Returns |
//! |----------|---------|---------|
//! | `GetSetting` | Get single registry value | `String` |
//! | `GetAllSettings` | Get all values in a section | `Variant` array |
//! | `SaveSetting` | Save registry value | N/A (statement) |
//! | `DeleteSetting` | Delete registry key/section | N/A (statement) |
//!
//! ## Platform Compatibility
//!
//! - **Windows**: Full support, uses `HKEY_CURRENT_USER` registry hive
//! - **Other Platforms**: May use alternative storage mechanisms or be unsupported
//! - **Registry Location**: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section\key`
//!
//! ## Limitations
//!
//! - Only accesses `HKEY_CURRENT_USER` hive (use Windows API for other hives)
//! - Returns `String` type only (requires conversion for other types)
//! - No direct way to check if a key exists (use unique default value trick)
//! - Limited to VB's registry structure (use Windows API for custom locations)
//! - No support for `REG_BINARY` or other complex registry types
//! - Settings are user-specific, not machine-wide
//!
//! ## Related Functions
//!
//! - `GetAllSettings`: Returns all key settings and their values from a registry section
//! - `SaveSetting`: Saves or creates an application entry in the Windows registry
//! - `DeleteSetting`: Deletes a section or key setting from the Windows registry
//! - `Environ`: Returns the string associated with an environment variable
//! - `Command`: Returns the argument portion of the command line
//! - `App.Path`: Returns the path where the application executable is located

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn getsetting_basic() {
        let source = r#"
Sub Test()
    value = GetSetting("MyApp", "Section", "Key", "Default")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_three_params() {
        let source = r#"
Sub Test()
    value = GetSetting("MyApp", "Section", "Key")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_in_function() {
        let source = r#"
Function GetUserName() As String
    GetUserName = GetSetting("MyApp", "User", "Name", "Guest")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_with_conversion() {
        let source = r#"
Sub Test()
    Dim x As Integer
    x = CInt(GetSetting("MyApp", "Settings", "Value", "0"))
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_if_statement() {
        let source = r#"
Sub Test()
    If GetSetting("MyApp", "Options", "AutoSave", "False") = "True" Then
        EnableAutoSave
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        files(i) = GetSetting("MyApp", "Recent", "File" & i, "")
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_select_case() {
        let source = r#"
Sub Test()
    Select Case GetSetting("MyApp", "UI", "Theme", "Default")
        Case "Dark"
            ApplyDarkTheme
        Case "Light"
            ApplyLightTheme
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_do_loop() {
        let source = r#"
Sub Test()
    Do While GetSetting("MyApp", "Status", "Running", "True") = "True"
        DoWork
    Loop
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_class_member() {
        let source = r#"
Private Sub Class_Initialize()
    m_setting = GetSetting("MyApp", "Config", "Value", "")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_type_field() {
        let source = r#"
Sub Test()
    Dim cfg As ConfigType
    cfg.server = GetSetting("MyApp", "Database", "Server", "localhost")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_collection_add() {
        let source = r#"
Sub Test()
    Dim col As New Collection
    col.Add GetSetting("MyApp", "Paths", "Data", App.Path)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_with_statement() {
        let source = r#"
Sub Test()
    With Form1
        .BackColor = CLng(GetSetting("MyApp", "Colors", "Background", "16777215"))
    End With
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print GetSetting("MyApp", "Debug", "Level", "Info")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_msgbox() {
        let source = r#"
Sub Test()
    MsgBox GetSetting("MyApp", "Messages", "Welcome", "Hello!")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_property() {
        let source = r#"
Property Get AppName() As String
    AppName = GetSetting("MyApp", "Info", "Name", "MyApplication")
End Property
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_concatenation() {
        let source = r#"
Sub Test()
    path = GetSetting("MyApp", "Paths", "Base", "C:\") & "Data"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_for_each() {
        let source = r#"
Sub Test()
    Dim ctrl As Control
    For Each ctrl In Controls
        ctrl.Tag = GetSetting("MyApp", "Controls", ctrl.Name, "")
    Next
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    value = GetSetting("MyApp", "Config", "Setting", "Default")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_comparison() {
        let source = r#"
Sub Test()
    If Len(GetSetting("MyApp", "User", "Name", "")) = 0 Then
        ShowWelcome
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_array_assignment() {
        let source = r#"
Sub Test()
    Dim settings(1 To 5) As String
    settings(1) = GetSetting("MyApp", "Config", "Setting1", "")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_function_argument() {
        let source = r#"
Sub Test()
    ProcessConfig GetSetting("MyApp", "Config", "File", "default.cfg")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_nested_call() {
        let source = r#"
Sub Test()
    value = UCase(GetSetting("MyApp", "Text", "Value", "default"))
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_iif() {
        let source = r#"
Sub Test()
    value = IIf(GetSetting("MyApp", "Options", "Mode", "") = "Advanced", 1, 0)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_form_load() {
        let source = r#"
Private Sub Form_Load()
    Me.Left = CLng(GetSetting("MyApp", "MainForm", "Left", "0"))
    Me.Top = CLng(GetSetting("MyApp", "MainForm", "Top", "0"))
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_multiple_assignments() {
        let source = r#"
Sub Test()
    server = GetSetting("MyApp", "Database", "Server", "localhost")
    database = GetSetting("MyApp", "Database", "Name", "MyDB")
    port = GetSetting("MyApp", "Database", "Port", "1433")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_parentheses() {
        let source = r#"
Sub Test()
    value = (GetSetting("MyApp", "Config", "Value", "0"))
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getsetting_string_builder() {
        let source = r#"
Sub Test()
    msg = "Server: " & GetSetting("MyApp", "DB", "Server", "localhost") & vbCrLf
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
