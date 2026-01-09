//! `GetAllSettings` Function
//!
//! Returns a list of key settings and their respective values (originally created with `SaveSetting`) from an application's entry in the Windows registry.
//!
//! # Syntax
//!
//! ```vb
//! GetAllSettings(appname, section)
//! ```
//!
//! # Parameters
//!
//! - `appname` - Required. String expression containing the name of the application or project whose key settings are requested.
//! - `section` - Required. String expression containing the name of the section whose key settings are requested.
//!
//! # Return Value
//!
//! Returns a `Variant` containing a two-dimensional array of strings. The first dimension contains the key names, and the second dimension contains the corresponding values.
//!
//! # Remarks
//!
//! - `GetAllSettings` returns an uninitialized `Variant` if either `appname` or `section` does not exist.
//! - The returned array is zero-based with two columns: column 0 contains key names, column 1 contains values.
//! - Works with the Windows registry (`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section`).
//! - On Windows, settings are stored in: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section`.
//! - Use `SaveSetting` to write values that `GetAllSettings` can retrieve.
//! - Use `GetSetting` to retrieve individual settings.
//! - Use `DeleteSetting` to remove settings from the registry.
//! - `GetAllSettings` is Windows-specific and relies on the registry.
//!
//! # Typical Uses
//!
//! - Loading application configuration settings
//! - Retrieving user preferences
//! - Reading saved window positions and sizes
//! - Loading multiple related settings at once
//! - Migrating settings between versions
//! - Exporting application configuration
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Retrieve all settings for a section
//! Dim allSettings As Variant
//! Dim i As Long
//!
//! allSettings = GetAllSettings("MyApp", "Preferences")
//!
//! If IsEmpty(allSettings) Then
//!     Debug.Print "No settings found"
//! Else
//!     For i = LBound(allSettings, 1) To UBound(allSettings, 1)
//!         Debug.Print allSettings(i, 0) & " = " & allSettings(i, 1)
//!     Next i
//! End If
//!
//! ' Check if settings exist
//! Dim settings As Variant
//! settings = GetAllSettings("MyApp", "WindowPosition")
//!
//! If Not IsEmpty(settings) Then
//!     ' Settings exist, process them
//!     MsgBox "Found " & (UBound(settings, 1) + 1) & " settings"
//! End If
//!
//! ' Load and apply settings
//! Dim appSettings As Variant
//! appSettings = GetAllSettings("MyApp", "Options")
//!
//! If Not IsEmpty(appSettings) Then
//!     Dim j As Long
//!     For j = 0 To UBound(appSettings, 1)
//!         Select Case appSettings(j, 0)
//!             Case "Theme"
//!                 ApplyTheme appSettings(j, 1)
//!             Case "Language"
//!                 SetLanguage appSettings(j, 1)
//!         End Select
//!     Next j
//! End If
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Load All Application Settings
//!
//! ```vb
//! Sub LoadApplicationSettings()
//!     Dim settings As Variant
//!     Dim i As Long
//!     
//!     settings = GetAllSettings("MyApp", "General")
//!     
//!     If IsEmpty(settings) Then
//!         ' No settings found, use defaults
//!         SetDefaultSettings
//!     Else
//!         For i = LBound(settings, 1) To UBound(settings, 1)
//!             ProcessSetting settings(i, 0), settings(i, 1)
//!         Next i
//!     End If
//! End Sub
//!
//! Sub ProcessSetting(keyName As String, value As String)
//!     Select Case keyName
//!         Case "AutoSave"
//!             chkAutoSave.Value = IIf(value = "True", 1, 0)
//!         Case "SaveInterval"
//!             txtInterval.Text = value
//!         Case "BackupEnabled"
//!             chkBackup.Value = IIf(value = "True", 1, 0)
//!     End Select
//! End Sub
//! ```
//!
//! ## 2. Restore Window Position
//!
//! ```vb
//! Sub RestoreWindowPosition(formName As Form)
//!     Dim settings As Variant
//!     Dim i As Long
//!     
//!     settings = GetAllSettings(App.Title, "WindowPos")
//!     
//!     If Not IsEmpty(settings) Then
//!         For i = LBound(settings, 1) To UBound(settings, 1)
//!             Select Case settings(i, 0)
//!                 Case "Left"
//!                     formName.Left = CLng(settings(i, 1))
//!                 Case "Top"
//!                     formName.Top = CLng(settings(i, 1))
//!                 Case "Width"
//!                     formName.Width = CLng(settings(i, 1))
//!                 Case "Height"
//!                     formName.Height = CLng(settings(i, 1))
//!                 Case "WindowState"
//!                     formName.WindowState = CInt(settings(i, 1))
//!             End Select
//!         Next i
//!     End If
//! End Sub
//! ```
//!
//! ## 3. Load User Preferences
//!
//! ```vb
//! Function LoadUserPreferences() As Collection
//!     Dim settings As Variant
//!     Dim prefs As New Collection
//!     Dim i As Long
//!     
//!     settings = GetAllSettings(App.Title, "UserPreferences")
//!     
//!     If Not IsEmpty(settings) Then
//!         For i = LBound(settings, 1) To UBound(settings, 1)
//!             prefs.Add settings(i, 1), settings(i, 0)
//!         Next i
//!     End If
//!     
//!     Set LoadUserPreferences = prefs
//! End Function
//!
//! ' Usage
//! Sub ApplyUserPreferences()
//!     Dim prefs As Collection
//!     Set prefs = LoadUserPreferences()
//!     
//!     If prefs.Count > 0 Then
//!         On Error Resume Next
//!         txtFontSize.Text = prefs("FontSize")
//!         cboTheme.Text = prefs("Theme")
//!         chkShowToolbar.Value = IIf(prefs("ShowToolbar") = "True", 1, 0)
//!         On Error GoTo 0
//!     End If
//! End Sub
//! ```
//!
//! ## 4. Export Settings to File
//!
//! ```vb
//! Sub ExportSettingsToFile(filename As String)
//!     Dim settings As Variant
//!     Dim fileNum As Integer
//!     Dim i As Long
//!     
//!     settings = GetAllSettings(App.Title, "Settings")
//!     
//!     If IsEmpty(settings) Then
//!         MsgBox "No settings to export"
//!         Exit Sub
//!     End If
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     Print #fileNum, "[Settings]"
//!     
//!     For i = LBound(settings, 1) To UBound(settings, 1)
//!         Print #fileNum, settings(i, 0) & "=" & settings(i, 1)
//!     Next i
//!     
//!     Close #fileNum
//!     MsgBox "Settings exported successfully"
//! End Sub
//! ```
//!
//! ## 5. Display Settings in `ListBox`
//!
//! ```vb
//! Sub PopulateSettingsList(lst As ListBox)
//!     Dim settings As Variant
//!     Dim i As Long
//!     
//!     lst.Clear
//!     settings = GetAllSettings(App.Title, "Configuration")
//!     
//!     If IsEmpty(settings) Then
//!         lst.AddItem "(No settings found)"
//!     Else
//!         For i = LBound(settings, 1) To UBound(settings, 1)
//!             lst.AddItem settings(i, 0) & " = " & settings(i, 1)
//!         Next i
//!     End If
//! End Sub
//! ```
//!
//! ## 6. Validate Settings
//!
//! ```vb
//! Function ValidateSettings(appName As String, section As String) As Boolean
//!     Dim settings As Variant
//!     Dim i As Long
//!     Dim isValid As Boolean
//!     
//!     settings = GetAllSettings(appName, section)
//!     
//!     If IsEmpty(settings) Then
//!         ValidateSettings = False
//!         Exit Function
//!     End If
//!     
//!     isValid = True
//!     
//!     For i = LBound(settings, 1) To UBound(settings, 1)
//!         ' Validate each setting
//!         If Not IsValidSetting(settings(i, 0), settings(i, 1)) Then
//!             Debug.Print "Invalid setting: " & settings(i, 0)
//!             isValid = False
//!         End If
//!     Next i
//!     
//!     ValidateSettings = isValid
//! End Function
//!
//! Function IsValidSetting(keyName As String, value As String) As Boolean
//!     ' Implement validation logic
//!     IsValidSetting = True
//!     
//!     Select Case keyName
//!         Case "Port"
//!             IsValidSetting = IsNumeric(value) And CLng(value) > 0 And CLng(value) < 65536
//!         Case "Timeout"
//!             IsValidSetting = IsNumeric(value) And CLng(value) > 0
//!         Case "Enabled"
//!             IsValidSetting = (value = "True" Or value = "False")
//!     End Select
//! End Function
//! ```
//!
//! ## 7. Compare Settings Between Sections
//!
//! ```vb
//! Sub CompareSettings(section1 As String, section2 As String)
//!     Dim settings1 As Variant
//!     Dim settings2 As Variant
//!     Dim i As Long
//!     
//!     settings1 = GetAllSettings(App.Title, section1)
//!     settings2 = GetAllSettings(App.Title, section2)
//!     
//!     Debug.Print "Comparing " & section1 & " vs " & section2
//!     Debug.Print String(50, "=")
//!     
//!     If IsEmpty(settings1) Then
//!         Debug.Print section1 & " has no settings"
//!     ElseIf IsEmpty(settings2) Then
//!         Debug.Print section2 & " has no settings"
//!     Else
//!         For i = LBound(settings1, 1) To UBound(settings1, 1)
//!             Debug.Print section1 & "." & settings1(i, 0) & " = " & settings1(i, 1)
//!         Next i
//!         
//!         Debug.Print ""
//!         
//!         For i = LBound(settings2, 1) To UBound(settings2, 1)
//!             Debug.Print section2 & "." & settings2(i, 0) & " = " & settings2(i, 1)
//!         Next i
//!     End If
//! End Sub
//! ```
//!
//! ## 8. Migrate Settings to New Version
//!
//! ```vb
//! Sub MigrateSettings(oldSection As String, newSection As String)
//!     Dim settings As Variant
//!     Dim i As Long
//!     
//!     ' Get all settings from old section
//!     settings = GetAllSettings(App.Title, oldSection)
//!     
//!     If IsEmpty(settings) Then
//!         Debug.Print "No settings to migrate"
//!         Exit Sub
//!     End If
//!     
//!     ' Save to new section
//!     For i = LBound(settings, 1) To UBound(settings, 1)
//!         SaveSetting App.Title, newSection, settings(i, 0), settings(i, 1)
//!     Next i
//!     
//!     Debug.Print "Migrated " & (UBound(settings, 1) + 1) & " settings"
//! End Sub
//! ```
//!
//! ## 9. Create Settings Dictionary
//!
//! ```vb
//! Function GetSettingsDictionary() As Object
//!     Dim settings As Variant
//!     Dim dict As Object
//!     Dim i As Long
//!     
//!     Set dict = CreateObject("Scripting.Dictionary")
//!     settings = GetAllSettings(App.Title, "Config")
//!     
//!     If Not IsEmpty(settings) Then
//!         For i = LBound(settings, 1) To UBound(settings, 1)
//!             dict.Add settings(i, 0), settings(i, 1)
//!         Next i
//!     End If
//!     
//!     Set GetSettingsDictionary = dict
//! End Function
//!
//! ' Usage
//! Sub UseSettingsDictionary()
//!     Dim settings As Object
//!     Set settings = GetSettingsDictionary()
//!     
//!     If settings.Exists("ServerURL") Then
//!         Debug.Print "Server: " & settings("ServerURL")
//!     End If
//! End Sub
//! ```
//!
//! ## 10. Bulk Settings Editor
//!
//! ```vb
//! Sub EditAllSettings(appName As String, section As String)
//!     Dim settings As Variant
//!     Dim i As Long
//!     Dim newValue As String
//!     
//!     settings = GetAllSettings(appName, section)
//!     
//!     If IsEmpty(settings) Then
//!         MsgBox "No settings found"
//!         Exit Sub
//!     End If
//!     
//!     For i = LBound(settings, 1) To UBound(settings, 1)
//!         newValue = InputBox("Enter new value for: " & settings(i, 0), _
//!                            "Edit Setting", settings(i, 1))
//!         
//!         If newValue <> "" Then
//!             SaveSetting appName, section, settings(i, 0), newValue
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. Settings Manager Class
//!
//! ```vb
//! ' Class: SettingsManager
//! Private m_AppName As String
//! Private m_Section As String
//! Private m_Settings As Variant
//!
//! Public Sub Initialize(appName As String, section As String)
//!     m_AppName = appName
//!     m_Section = section
//!     RefreshSettings
//! End Sub
//!
//! Public Sub RefreshSettings()
//!     m_Settings = GetAllSettings(m_AppName, m_Section)
//! End Sub
//!
//! Public Function GetValue(keyName As String, _
//!                          Optional defaultValue As String = "") As String
//!     Dim i As Long
//!     
//!     If IsEmpty(m_Settings) Then
//!         GetValue = defaultValue
//!         Exit Function
//!     End If
//!     
//!     For i = LBound(m_Settings, 1) To UBound(m_Settings, 1)
//!         If m_Settings(i, 0) = keyName Then
//!             GetValue = m_Settings(i, 1)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetValue = defaultValue
//! End Function
//!
//! Public Function SettingExists(keyName As String) As Boolean
//!     Dim i As Long
//!     
//!     If IsEmpty(m_Settings) Then
//!         SettingExists = False
//!         Exit Function
//!     End If
//!     
//!     For i = LBound(m_Settings, 1) To UBound(m_Settings, 1)
//!         If m_Settings(i, 0) = keyName Then
//!             SettingExists = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     SettingExists = False
//! End Function
//!
//! Public Property Get SettingCount() As Long
//!     If IsEmpty(m_Settings) Then
//!         SettingCount = 0
//!     Else
//!         SettingCount = UBound(m_Settings, 1) + 1
//!     End If
//! End Property
//! ```
//!
//! ## 2. Settings Backup and Restore
//!
//! ```vb
//! Type SettingsBackup
//!     AppName As String
//!     Section As String
//!     Settings As Variant
//!     BackupDate As Date
//! End Type
//!
//! Function BackupSettings(appName As String, section As String) As SettingsBackup
//!     Dim backup As SettingsBackup
//!     
//!     backup.AppName = appName
//!     backup.Section = section
//!     backup.Settings = GetAllSettings(appName, section)
//!     backup.BackupDate = Now
//!     
//!     BackupSettings = backup
//! End Function
//!
//! Sub RestoreSettings(backup As SettingsBackup)
//!     Dim i As Long
//!     
//!     If IsEmpty(backup.Settings) Then
//!         MsgBox "No settings to restore"
//!         Exit Sub
//!     End If
//!     
//!     ' Clear existing settings first
//!     DeleteSetting backup.AppName, backup.Section
//!     
//!     ' Restore backed up settings
//!     For i = LBound(backup.Settings, 1) To UBound(backup.Settings, 1)
//!         SaveSetting backup.AppName, backup.Section, _
//!                    backup.Settings(i, 0), backup.Settings(i, 1)
//!     Next i
//!     
//!     MsgBox "Settings restored from " & Format(backup.BackupDate, "yyyy-mm-dd hh:nn:ss")
//! End Sub
//! ```
//!
//! ## 3. Settings Encryption/Decryption
//!
//! ```vb
//! Function GetEncryptedSettings(appName As String, section As String, _
//!                               password As String) As Variant
//!     Dim settings As Variant
//!     Dim decrypted() As String
//!     Dim i As Long
//!     
//!     settings = GetAllSettings(appName, section)
//!     
//!     If IsEmpty(settings) Then
//!         GetEncryptedSettings = Empty
//!         Exit Function
//!     End If
//!     
//!     ReDim decrypted(LBound(settings, 1) To UBound(settings, 1), 0 To 1)
//!     
//!     For i = LBound(settings, 1) To UBound(settings, 1)
//!         decrypted(i, 0) = settings(i, 0)
//!         decrypted(i, 1) = DecryptString(settings(i, 1), password)
//!     Next i
//!     
//!     GetEncryptedSettings = decrypted
//! End Function
//!
//! Function DecryptString(encrypted As String, password As String) As String
//!     ' Simple XOR encryption for demonstration
//!     Dim i As Long
//!     Dim result As String
//!     Dim keyChar As Integer
//!     
//!     result = ""
//!     
//!     For i = 1 To Len(encrypted)
//!         keyChar = Asc(Mid(password, ((i - 1) Mod Len(password)) + 1, 1))
//!         result = result & Chr(Asc(Mid(encrypted, i, 1)) Xor keyChar)
//!     Next i
//!     
//!     DecryptString = result
//! End Function
//! ```
//!
//! ## 4. Multi-Section Settings Loader
//!
//! ```vb
//! Function LoadMultipleSections(appName As String, _
//!                               sections() As String) As Collection
//!     Dim allSettings As New Collection
//!     Dim i As Long
//!     Dim sectionSettings As Variant
//!     
//!     For i = LBound(sections) To UBound(sections)
//!         sectionSettings = GetAllSettings(appName, sections(i))
//!         
//!         If Not IsEmpty(sectionSettings) Then
//!             allSettings.Add sectionSettings, sections(i)
//!         End If
//!     Next i
//!     
//!     Set LoadMultipleSections = allSettings
//! End Function
//!
//! ' Usage
//! Sub LoadAllAppSettings()
//!     Dim sections() As String
//!     Dim allSettings As Collection
//!     Dim section As Variant
//!     
//!     sections = Split("General,Display,Network,Security", ",")
//!     Set allSettings = LoadMultipleSections(App.Title, sections)
//!     
//!     For Each section In allSettings
//!         Debug.Print "Section has " & (UBound(section, 1) + 1) & " settings"
//!     Next
//! End Sub
//! ```
//!
//! ## 5. Settings Change Detection
//!
//! ```vb
//! Type SettingsSnapshot
//!     Settings As Variant
//!     Timestamp As Date
//! End Type
//!
//! Private m_LastSnapshot As SettingsSnapshot
//!
//! Function TakeSnapshot(appName As String, section As String) As SettingsSnapshot
//!     Dim snapshot As SettingsSnapshot
//!     
//!     snapshot.Settings = GetAllSettings(appName, section)
//!     snapshot.Timestamp = Now
//!     
//!     TakeSnapshot = snapshot
//! End Function
//!
//! Function DetectChanges(appName As String, section As String) As Boolean
//!     Dim currentSettings As Variant
//!     Dim i As Long
//!     Dim changed As Boolean
//!     
//!     currentSettings = GetAllSettings(appName, section)
//!     
//!     ' Check if both are empty
//!     If IsEmpty(m_LastSnapshot.Settings) And IsEmpty(currentSettings) Then
//!         DetectChanges = False
//!         Exit Function
//!     End If
//!     
//!     ' Check if one is empty
//!     If IsEmpty(m_LastSnapshot.Settings) Or IsEmpty(currentSettings) Then
//!         DetectChanges = True
//!         Exit Function
//!     End If
//!     
//!     ' Check if different sizes
//!     If UBound(m_LastSnapshot.Settings, 1) <> UBound(currentSettings, 1) Then
//!         DetectChanges = True
//!         Exit Function
//!     End If
//!     
//!     ' Compare values
//!     changed = False
//!     For i = LBound(currentSettings, 1) To UBound(currentSettings, 1)
//!         If currentSettings(i, 0) <> m_LastSnapshot.Settings(i, 0) Or _
//!            currentSettings(i, 1) <> m_LastSnapshot.Settings(i, 1) Then
//!             changed = True
//!             Exit For
//!         End If
//!     Next i
//!     
//!     DetectChanges = changed
//! End Function
//! ```
//!
//! ## 6. Settings Validation Framework
//!
//! ```vb
//! Type ValidationRule
//!     KeyName As String
//!     DataType As String  ' "String", "Integer", "Boolean", "Date"
//!     MinValue As Variant
//!     MaxValue As Variant
//!     Required As Boolean
//! End Type
//!
//! Function ValidateAllSettings(appName As String, section As String, _
//!                              rules() As ValidationRule) As Collection
//!     Dim settings As Variant
//!     Dim errors As New Collection
//!     Dim i As Long, j As Long
//!     Dim found As Boolean
//!     Dim settingValue As String
//!     
//!     settings = GetAllSettings(appName, section)
//!     
//!     ' Check each rule
//!     For i = LBound(rules) To UBound(rules)
//!         found = False
//!         
//!         If Not IsEmpty(settings) Then
//!             For j = LBound(settings, 1) To UBound(settings, 1)
//!                 If settings(j, 0) = rules(i).KeyName Then
//!                     found = True
//!                     settingValue = settings(j, 1)
//!                     
//!                     ' Validate data type and range
//!                     If Not ValidateValue(settingValue, rules(i)) Then
//!                         errors.Add "Invalid value for " & rules(i).KeyName
//!                     End If
//!                     Exit For
//!                 End If
//!             Next j
//!         End If
//!         
//!         If Not found And rules(i).Required Then
//!             errors.Add "Missing required setting: " & rules(i).KeyName
//!         End If
//!     Next i
//!     
//!     Set ValidateAllSettings = errors
//! End Function
//!
//! Function ValidateValue(value As String, rule As ValidationRule) As Boolean
//!     Select Case rule.DataType
//!         Case "Integer"
//!             If Not IsNumeric(value) Then
//!                 ValidateValue = False
//!             Else
//!                 Dim intVal As Long
//!                 intVal = CLng(value)
//!                 ValidateValue = (intVal >= rule.MinValue And intVal <= rule.MaxValue)
//!             End If
//!         Case "Boolean"
//!             ValidateValue = (value = "True" Or value = "False")
//!         Case Else
//!             ValidateValue = True
//!     End Select
//! End Function
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeGetAllSettings(appName As String, section As String) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     Dim settings As Variant
//!     settings = GetAllSettings(appName, section)
//!     
//!     SafeGetAllSettings = settings
//!     Exit Function
//!     
//! ErrorHandler:
//!     Debug.Print "Error retrieving settings: " & Err.Description
//!     SafeGetAllSettings = Empty
//! End Function
//! ```
//!
//! Common issues:
//! - **Empty return value**: Section or application does not exist in registry
//! - **Registry access denied**: Insufficient permissions to read registry
//! - **Invalid section name**: Special characters or invalid path
//!
//! # Performance Considerations
//!
//! - `GetAllSettings` reads from the Windows registry, which is relatively fast
//! - For frequently accessed settings, consider caching the results
//! - Reading all settings at once is more efficient than multiple `GetSetting` calls
//! - Registry access can be affected by antivirus software
//! - Consider using INI files or XML for cross-platform compatibility
//!
//! # Best Practices
//!
//! 1. **Always check for Empty** before using the returned array
//! 2. **Use meaningful section names** to organize settings logically
//! 3. **Cache settings** in memory if accessed frequently
//! 4. **Validate settings** after retrieval
//! 5. **Provide defaults** when settings don't exist
//! 6. **Document registry structure** for your application
//! 7. **Consider cleanup** - use `DeleteSetting` when settings are no longer needed
//!
//! # Comparison with Other Functions
//!
//! ## `GetAllSettings` vs `GetSetting`
//!
//! ```vb
//! ' GetAllSettings - Retrieve all settings at once
//! Dim allSettings As Variant
//! allSettings = GetAllSettings("MyApp", "Config")
//!
//! ' GetSetting - Retrieve one setting at a time
//! Dim value As String
//! value = GetSetting("MyApp", "Config", "Theme", "Default")
//! ```
//!
//! ## `GetAllSettings` vs File-Based Storage
//!
//! ```vb
//! ' GetAllSettings - Windows registry
//! settings = GetAllSettings(App.Title, "Settings")
//!
//! ' File-based - INI file or XML (more portable)
//! ' Requires custom parsing code
//! ```
//!
//! # Limitations
//!
//! - Windows-specific (uses Windows registry)
//! - Limited to `HKEY_CURRENT_USER` hive
//! - `String` values only (need to parse numbers, dates, etc.)
//! - Registry size limits (though rarely hit in practice)
//! - No built-in encryption or security
//! - Requires appropriate registry permissions
//! - Not suitable for large amounts of data
//! - Two-dimensional array only (fixed structure)
//!
//! # Registry Location
//!
//! Settings are stored at:
//! ```text
//! HKEY_CURRENT_USER\Software\VB and VBA Program Settings\appname\section
//! ```
//!
//! # Related Functions
//!
//! - `GetSetting` - Returns a single key setting from the registry
//! - `SaveSetting` - Saves or creates an application entry in the registry
//! - `DeleteSetting` - Deletes a section or key setting from the registry
//! - `Environ` - Returns the string associated with an operating system environment variable
//! - `Command` - Returns the command-line arguments

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn getallsettings_basic() {
        let source = r#"allSettings = GetAllSettings("MyApp", "Preferences")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_with_variable() {
        let source = r"settings = GetAllSettings(appName, sectionName)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_app_title() {
        let source = r#"settings = GetAllSettings(App.Title, "WindowPos")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_if_check() {
        let source = r#"If IsEmpty(GetAllSettings("MyApp", "Config")) Then MsgBox "No settings""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_in_function() {
        let source = r#"Function LoadSettings() As Variant
    LoadSettings = GetAllSettings(App.Title, "General")
End Function"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_for_loop() {
        let source = r"For i = LBound(settings, 1) To UBound(settings, 1)
    Debug.Print settings(i, 0)
Next i";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_array_access() {
        let source = r#"value = allSettings(i, 0) & " = " & allSettings(i, 1)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_select_case() {
        let source = r#"Select Case allSettings(j, 0)
    Case "Theme"
        ApplyTheme allSettings(j, 1)
End Select"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_dim_statement() {
        let source = r#"Dim settings As Variant
settings = GetAllSettings("MyApp", "Settings")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_not_isempty() {
        let source = r"If Not IsEmpty(GetAllSettings(appName, section)) Then ProcessSettings";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_debug_print() {
        let source =
            r#"Debug.Print "Settings count: " & (UBound(GetAllSettings(appName, section), 1) + 1)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_msgbox() {
        let source =
            r#"MsgBox "Found " & (UBound(GetAllSettings("App", "Section"), 1) + 1) & " settings""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_error_handling() {
        let source = r"On Error GoTo ErrorHandler
settings = GetAllSettings(appName, section)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_listbox() {
        let source = r#"lst.AddItem allSettings(i, 0) & " = " & allSettings(i, 1)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_collection_add() {
        let source = r"prefs.Add settings(i, 1), settings(i, 0)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_file_export() {
        let source = r#"Print #fileNum, settings(i, 0) & "=" & settings(i, 1)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_comparison() {
        let source = r"If settings1(i, 0) <> settings2(i, 0) Then changed = True";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_dictionary() {
        let source = r"dict.Add settings(i, 0), settings(i, 1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_inputbox() {
        let source = r#"newValue = InputBox("Enter new value for: " & settings(i, 0), "Edit Setting", settings(i, 1))"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_class_member() {
        let source = r"m_Settings = GetAllSettings(m_AppName, m_Section)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_type_field() {
        let source = r"backup.Settings = GetAllSettings(appName, section)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_do_loop() {
        let source = r"Do While i <= UBound(settings, 1)
    ProcessSetting settings(i, 0), settings(i, 1)
    i = i + 1
Loop";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_concatenation() {
        let source = r#"msg = "Key: " & settings(i, 0) & " Value: " & settings(i, 1)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_redim() {
        let source = r"ReDim decrypted(LBound(settings, 1) To UBound(settings, 1), 0 To 1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_savesetting() {
        let source = r"SaveSetting appName, section, settings(i, 0), settings(i, 1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getallsettings_property() {
        let source = r"If IsEmpty(m_Settings) Then SettingCount = 0";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/environment/getallsettings",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
