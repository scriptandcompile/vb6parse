//! # LoadResString Function
//!
//! Returns a string from a resource (.res) file.
//!
//! ## Syntax
//!
//! ```vb
//! LoadResString(index)
//! ```
//!
//! ## Parameters
//!
//! - `index` (Required): Integer identifying the string resource
//!   - Must be a numeric ID (string names not supported for string resources)
//!   - Must match the ID used when the resource was compiled
//!   - Typically ranges from 1 to 65535
//!
//! ## Return Value
//!
//! Returns a String:
//! - String containing the text from the resource file
//! - Empty string ("") if resource not found (in some VB versions)
//! - May raise error 32813 if resource not found
//! - Preserves all formatting including line breaks
//! - Unicode strings supported in VB6
//!
//! ## Remarks
//!
//! The LoadResString function loads text from embedded resources:
//!
//! - Loads strings from compiled resource (.res) files
//! - Resource file must be linked to project at compile time
//! - Primary method for internationalization (i18n) in VB6
//! - Allows localizing applications without code changes
//! - Strings can be translated by replacing resource file
//! - No external text files needed at runtime
//! - Embedded in compiled EXE/DLL
//! - Only one resource file per project
//! - Resource file added via Project > Add File
//! - Resource files created with Resource Editor or RC.EXE
//! - Index must be numeric (string names not supported)
//! - Common for error messages, prompts, labels
//! - Supports Unicode in VB6
//! - Error 32813: "Resource not found" if ID doesn't exist
//! - Error 48: "Error loading from file" if resource file corrupt
//! - More maintainable than hardcoded strings
//! - Easier to update text without recompiling code
//! - Standard practice for multi-language applications
//! - Can store long text passages
//! - Supports special characters and formatting
//!
//! ## Typical Uses
//!
//! 1. **Load Error Message**
//!    ```vb
//!    MsgBox LoadResString(1001), vbCritical
//!    ```
//!
//! 2. **Load Form Caption**
//!    ```vb
//!    Me.Caption = LoadResString(2001)
//!    ```
//!
//! 3. **Load Label Text**
//!    ```vb
//!    lblWelcome.Caption = LoadResString(3001)
//!    ```
//!
//! 4. **Load Menu Caption**
//!    ```vb
//!    mnuFile.Caption = LoadResString(4001)
//!    ```
//!
//! 5. **Load Button Caption**
//!    ```vb
//!    cmdOK.Caption = LoadResString(5001)
//!    ```
//!
//! 6. **Load MessageBox Text**
//!    ```vb
//!    MsgBox LoadResString(6001), vbInformation
//!    ```
//!
//! 7. **Load StatusBar Text**
//!    ```vb
//!    StatusBar1.SimpleText = LoadResString(7001)
//!    ```
//!
//! 8. **Load ToolTip Text**
//!    ```vb
//!    cmdSave.ToolTipText = LoadResString(8001)
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Loading Messages
//! ```vb
//! ' Load various UI strings from resources
//! Me.Caption = LoadResString(1001)          ' "My Application"
//! lblTitle.Caption = LoadResString(1002)    ' "Welcome!"
//! cmdOK.Caption = LoadResString(1003)       ' "OK"
//! cmdCancel.Caption = LoadResString(1004)   ' "Cancel"
//! ```
//!
//! ### Example 2: Error Messages
//! ```vb
//! ' Use resource strings for error messages
//! If Not fileExists Then
//!     MsgBox LoadResString(2001), vbCritical  ' "File not found"
//! End If
//!
//! If accessDenied Then
//!     MsgBox LoadResString(2002), vbCritical  ' "Access denied"
//! End If
//! ```
//!
//! ### Example 3: Error Handling
//! ```vb
//! On Error Resume Next
//! Dim msg As String
//! msg = LoadResString(9999)
//! If Err.Number = 32813 Then
//!     msg = "String resource not found!"
//!     Err.Clear
//! End If
//! MsgBox msg
//! ```
//!
//! ### Example 4: Form Initialization
//! ```vb
//! Private Sub Form_Load()
//!     ' Load all UI strings from resources
//!     Me.Caption = LoadResString(1001)
//!     lblName.Caption = LoadResString(1002)
//!     lblAddress.Caption = LoadResString(1003)
//!     cmdSave.Caption = LoadResString(1004)
//!     cmdCancel.Caption = LoadResString(1005)
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: SafeLoadResString
//! ```vb
//! Function SafeLoadResString(ByVal resID As Integer, _
//!                            Optional ByVal defaultText As String = "") As String
//!     On Error Resume Next
//!     SafeLoadResString = LoadResString(resID)
//!     If Err.Number <> 0 Then
//!         SafeLoadResString = defaultText
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: LoadFormStrings
//! ```vb
//! Sub LoadFormStrings(frm As Form, ByVal baseID As Integer)
//!     Dim ctrl As Control
//!     Dim id As Integer
//!     
//!     On Error Resume Next
//!     frm.Caption = LoadResString(baseID)
//!     
//!     id = baseID + 1
//!     For Each ctrl In frm.Controls
//!         If TypeOf ctrl Is Label Or TypeOf ctrl Is CommandButton Then
//!             ctrl.Caption = LoadResString(id)
//!             id = id + 1
//!         End If
//!     Next ctrl
//! End Sub
//! ```
//!
//! ### Pattern 3: FormatResString
//! ```vb
//! Function FormatResString(ByVal resID As Integer, _
//!                          ParamArray args()) As String
//!     Dim template As String
//!     Dim i As Long
//!     
//!     template = LoadResString(resID)
//!     
//!     For i = LBound(args) To UBound(args)
//!         template = Replace(template, "{" & i & "}", CStr(args(i)))
//!     Next i
//!     
//!     FormatResString = template
//! End Function
//! ```
//!
//! ### Pattern 4: GetErrorMessage
//! ```vb
//! Function GetErrorMessage(ByVal errorCode As Long) As String
//!     Const BASE_ERROR_ID = 10000
//!     On Error Resume Next
//!     
//!     GetErrorMessage = LoadResString(BASE_ERROR_ID + errorCode)
//!     If Err.Number <> 0 Then
//!         GetErrorMessage = "Unknown error: " & errorCode
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: LoadMenuStrings
//! ```vb
//! Sub LoadMenuStrings()
//!     Const MENU_BASE = 4000
//!     
//!     mnuFile.Caption = LoadResString(MENU_BASE + 1)      ' "&File"
//!     mnuFileNew.Caption = LoadResString(MENU_BASE + 2)   ' "&New"
//!     mnuFileOpen.Caption = LoadResString(MENU_BASE + 3)  ' "&Open"
//!     mnuFileSave.Caption = LoadResString(MENU_BASE + 4)  ' "&Save"
//!     mnuFileExit.Caption = LoadResString(MENU_BASE + 5)  ' "E&xit"
//! End Sub
//! ```
//!
//! ### Pattern 6: CachedResString
//! ```vb
//! Dim resStringCache As New Collection
//!
//! Function CachedLoadResString(ByVal resID As Integer) As String
//!     Dim key As String
//!     On Error Resume Next
//!     
//!     key = "RES_" & resID
//!     CachedLoadResString = resStringCache(key)
//!     
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         CachedLoadResString = LoadResString(resID)
//!         resStringCache.Add CachedLoadResString, key
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: ResStringExists
//! ```vb
//! Function ResStringExists(ByVal resID As Integer) As Boolean
//!     On Error Resume Next
//!     Dim s As String
//!     s = LoadResString(resID)
//!     ResStringExists = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 8: LoadResStringArray
//! ```vb
//! Function LoadResStringArray(ByVal startID As Integer, _
//!                             ByVal count As Integer) As String()
//!     Dim result() As String
//!     Dim i As Integer
//!     
//!     ReDim result(0 To count - 1)
//!     
//!     On Error Resume Next
//!     For i = 0 To count - 1
//!         result(i) = LoadResString(startID + i)
//!         If Err.Number <> 0 Then
//!             result(i) = ""
//!             Err.Clear
//!         End If
//!     Next i
//!     
//!     LoadResStringArray = result
//! End Function
//! ```
//!
//! ### Pattern 9: ShowResMessage
//! ```vb
//! Sub ShowResMessage(ByVal resID As Integer, _
//!                    Optional ByVal icon As VbMsgBoxStyle = vbInformation)
//!     On Error Resume Next
//!     Dim msg As String
//!     msg = LoadResString(resID)
//!     
//!     If Err.Number = 0 Then
//!         MsgBox msg, icon
//!     Else
//!         MsgBox "Message resource " & resID & " not found", vbCritical
//!         Err.Clear
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 10: MultiLineResString
//! ```vb
//! Function MultiLineResString(ByVal resID As Integer) As String
//!     Dim text As String
//!     text = LoadResString(resID)
//!     ' Resource strings preserve line breaks
//!     MultiLineResString = text
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Localization Manager
//! ```vb
//! ' Class: LocalizationManager
//! Private m_cache As Collection
//! Private m_languageID As Integer
//!
//! Private Sub Class_Initialize()
//!     Set m_cache = New Collection
//!     m_languageID = 1033 ' Default to English (US)
//! End Sub
//!
//! Public Property Let LanguageID(ByVal newLanguage As Integer)
//!     m_languageID = newLanguage
//!     ClearCache
//! End Property
//!
//! Public Function GetString(ByVal baseID As Integer) As String
//!     Dim resID As Integer
//!     Dim key As String
//!     
//!     On Error Resume Next
//!     resID = baseID + m_languageID
//!     key = "STR_" & resID
//!     
//!     GetString = m_cache(key)
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         GetString = LoadResString(resID)
//!         If Err.Number = 0 Then
//!             m_cache.Add GetString, key
//!         Else
//!             ' Fallback to default language
//!             GetString = LoadResString(baseID + 1033)
//!             Err.Clear
//!         End If
//!     End If
//! End Function
//!
//! Public Sub LocalizeForm(frm As Form)
//!     Dim ctrl As Control
//!     On Error Resume Next
//!     
//!     ' Load form caption
//!     frm.Caption = GetString(GetFormBaseID(frm))
//!     
//!     ' Load control captions
//!     For Each ctrl In frm.Controls
//!         If HasCaption(ctrl) Then
//!             ctrl.Caption = GetString(GetControlID(ctrl))
//!         End If
//!     Next ctrl
//! End Sub
//!
//! Private Sub ClearCache()
//!     Set m_cache = New Collection
//! End Sub
//!
//! Private Function HasCaption(ctrl As Control) As Boolean
//!     HasCaption = TypeOf ctrl Is Label Or _
//!                  TypeOf ctrl Is CommandButton Or _
//!                  TypeOf ctrl Is CheckBox Or _
//!                  TypeOf ctrl Is OptionButton
//! End Function
//!
//! Private Sub Class_Terminate()
//!     Set m_cache = Nothing
//! End Sub
//! ```
//!
//! ### Example 2: Error Message System
//! ```vb
//! ' Module: ErrorMessages
//! Private Const ERR_BASE = 20000
//!
//! Public Enum AppError
//!     errFileNotFound = 1
//!     errAccessDenied = 2
//!     errInvalidFormat = 3
//!     errNetworkError = 4
//!     errDatabaseError = 5
//! End Enum
//!
//! Public Sub ShowError(ByVal errorType As AppError, _
//!                      Optional ByVal additionalInfo As String = "")
//!     Dim msg As String
//!     On Error Resume Next
//!     
//!     msg = LoadResString(ERR_BASE + errorType)
//!     If Err.Number <> 0 Then
//!         msg = "Unknown error occurred"
//!         Err.Clear
//!     End If
//!     
//!     If Len(additionalInfo) > 0 Then
//!         msg = msg & vbCrLf & vbCrLf & additionalInfo
//!     End If
//!     
//!     MsgBox msg, vbCritical, LoadResString(ERR_BASE)
//! End Sub
//!
//! Public Function GetErrorText(ByVal errorType As AppError) As String
//!     On Error Resume Next
//!     GetErrorText = LoadResString(ERR_BASE + errorType)
//!     If Err.Number <> 0 Then
//!         GetErrorText = "Unknown error"
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Multi-Language Application
//! ```vb
//! ' Form with language selection
//! Public Enum Language
//!     langEnglish = 0
//!     langSpanish = 1000
//!     langFrench = 2000
//!     langGerman = 3000
//! End Enum
//!
//! Private currentLanguage As Language
//!
//! Private Sub Form_Load()
//!     ' Default to English
//!     currentLanguage = langEnglish
//!     LoadLanguage
//! End Sub
//!
//! Private Sub cboLanguage_Click()
//!     Select Case cboLanguage.ListIndex
//!         Case 0: currentLanguage = langEnglish
//!         Case 1: currentLanguage = langSpanish
//!         Case 2: currentLanguage = langFrench
//!         Case 3: currentLanguage = langGerman
//!     End Select
//!     LoadLanguage
//! End Sub
//!
//! Private Sub LoadLanguage()
//!     Dim baseID As Integer
//!     baseID = 10000 + currentLanguage
//!     
//!     On Error Resume Next
//!     Me.Caption = LoadResString(baseID + 1)
//!     lblWelcome.Caption = LoadResString(baseID + 2)
//!     lblInstructions.Caption = LoadResString(baseID + 3)
//!     cmdStart.Caption = LoadResString(baseID + 4)
//!     cmdExit.Caption = LoadResString(baseID + 5)
//!     
//!     ' Update menu
//!     mnuFile.Caption = LoadResString(baseID + 10)
//!     mnuHelp.Caption = LoadResString(baseID + 11)
//! End Sub
//! ```
//!
//! ### Example 4: String Template System
//! ```vb
//! ' Module: StringTemplates
//! Private Const TEMPLATE_BASE = 30000
//!
//! Public Function GetFormattedString(ByVal templateID As Integer, _
//!                                    ParamArray values()) As String
//!     Dim template As String
//!     Dim result As String
//!     Dim i As Long
//!     
//!     On Error Resume Next
//!     template = LoadResString(TEMPLATE_BASE + templateID)
//!     If Err.Number <> 0 Then
//!         GetFormattedString = ""
//!         Err.Clear
//!         Exit Function
//!     End If
//!     
//!     result = template
//!     For i = LBound(values) To UBound(values)
//!         result = Replace(result, "{" & i & "}", CStr(values(i)))
//!     Next i
//!     
//!     GetFormattedString = result
//! End Function
//!
//! Public Function GetWelcomeMessage(ByVal userName As String) As String
//!     ' Template: "Welcome, {0}! You have {1} new messages."
//!     GetWelcomeMessage = GetFormattedString(1, userName, GetMessageCount())
//! End Function
//!
//! Public Function GetSaveConfirmation(ByVal filename As String) As String
//!     ' Template: "Do you want to save changes to {0}?"
//!     GetSaveConfirmation = GetFormattedString(2, filename)
//! End Function
//!
//! Private Function GetMessageCount() As Long
//!     ' Implementation would return actual message count
//!     GetMessageCount = 5
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 32813: Resource not found
//! On Error Resume Next
//! Dim msg As String
//! msg = LoadResString(9999)
//! If Err.Number = 32813 Then
//!     MsgBox "String resource not found!"
//! End If
//!
//! ' Error 48: Error loading from file
//! msg = LoadResString(1001)
//! If Err.Number = 48 Then
//!     MsgBox "Resource file is corrupt or missing!"
//! End If
//!
//! ' Safe loading pattern
//! Function TryLoadResString(ByVal resID As Integer, _
//!                           ByRef outString As String) As Boolean
//!     On Error Resume Next
//!     outString = LoadResString(resID)
//!     TryLoadResString = (Err.Number = 0)
//!     If Err.Number <> 0 Then
//!         outString = ""
//!     End If
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Access**: Strings embedded in EXE (very fast loading)
//! - **No File I/O**: No disk access required
//! - **No Caching**: Each call loads fresh copy (implement caching if needed)
//! - **Memory Efficient**: Strings only loaded when accessed
//! - **Cache Strategy**: For frequently used strings, cache in Collection or array
//! - **Startup Time**: Loading many strings at startup may slow Form_Load
//!
//! ## Best Practices
//!
//! 1. **Use constants** for string IDs for maintainability
//! 2. **Group by category** using ID ranges (1000s for errors, 2000s for menus, etc.)
//! 3. **Cache frequently used strings** to improve performance
//! 4. **Always handle errors** - resource might not exist
//! 5. **Document string IDs** in code or separate file
//! 6. **Use templates** with placeholders for dynamic content
//! 7. **Organize by language** using ID offsets (English: 0, Spanish: +1000, etc.)
//! 8. **Test all languages** before deployment
//! 9. **Provide fallbacks** for missing strings
//! 10. **Keep strings updated** in sync with code changes
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Return Type | Data Type |
//! |----------|---------|-------------|-----------|
//! | **LoadResString** | Load string from resources | String | Text strings |
//! | **LoadResPicture** | Load image from resources | StdPicture | Images |
//! | **LoadResData** | Load binary data from resources | Byte array | Binary data |
//! | **LoadString** (API) | Windows API alternative | String | Text strings |
//!
//! ## LoadResString vs Hardcoded Strings
//!
//! ```vb
//! ' Hardcoded - difficult to localize
//! MsgBox "File not found", vbCritical
//!
//! ' Resource string - easy to localize
//! MsgBox LoadResString(1001), vbCritical
//! ```
//!
//! **Advantages of LoadResString:**
//! - Easy localization (just replace .res file)
//! - Centralized string management
//! - No code changes needed for translations
//! - Consistent messaging across application
//!
//! ## String ID Organization
//!
//! ```vb
//! ' Recommended ID ranges
//! Const STR_APP_BASE = 1000         ' Application strings
//! Const STR_ERROR_BASE = 2000       ' Error messages
//! Const STR_MENU_BASE = 3000        ' Menu items
//! Const STR_DIALOG_BASE = 4000      ' Dialog messages
//! Const STR_STATUS_BASE = 5000      ' Status messages
//! Const STR_HELP_BASE = 6000        ' Help text
//!
//! ' Language offsets
//! Const LANG_ENGLISH = 0
//! Const LANG_SPANISH = 10000
//! Const LANG_FRENCH = 20000
//! ```
//!
//! ## Platform Notes
//!
//! - Available in VB6 (not in early VB versions)
//! - Requires resource file (.res) linked to project
//! - Resource file created with Resource Editor or RC.EXE
//! - Only one resource file per project
//! - Resources embedded in compiled EXE/DLL
//! - Supports Unicode strings in VB6
//! - Index must be Integer (1-65535)
//! - String names not supported (numeric IDs only)
//! - Standard method for internationalization
//! - Preserves formatting including line breaks
//!
//! ## Limitations
//!
//! - **One Resource File**: Only one .res file per project
//! - **Numeric IDs Only**: Cannot use string names for string resources
//! - **Compile Time**: Must recompile to update strings
//! - **No Modification**: Cannot modify resources at runtime
//! - **Limited Editor**: VB6 Resource Editor is basic
//! - **ID Range**: Limited to 1-65535
//! - **No Encryption**: Strings easily extractable from EXE
//! - **No Formatting**: No printf-style formatting (must implement manually)
//! - **No Pluralization**: No built-in plural form handling
//! - **No Context**: All strings in flat namespace
//!
//! ## Related Functions
//!
//! - `LoadResPicture`: Load picture from resource file
//! - `LoadResData`: Load binary data from resource file
//! - `LoadPicture`: Load picture from external file
//! - `Format`: Format strings with values
//! - `Replace`: Replace placeholders in strings

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_loadresstring_basic() {
        let source = r#"
            Dim msg As String
            msg = LoadResString(1001)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_msgbox() {
        let source = r#"
            MsgBox LoadResString(2001), vbCritical
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_if_statement() {
        let source = r#"
            If Not fileExists Then
                MsgBox LoadResString(3001), vbCritical
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_caption() {
        let source = r#"
            Me.Caption = LoadResString(4001)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_form_load() {
        let source = r#"
            Private Sub Form_Load()
                lblWelcome.Caption = LoadResString(5001)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_for_loop() {
        let source = r#"
            For i = 1 To 5
                labels(i).Caption = LoadResString(6000 + i)
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_function_return() {
        let source = r#"
            Function GetErrorMessage() As String
                GetErrorMessage = LoadResString(7001)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_error_handling() {
        let source = r#"
            On Error Resume Next
            msg = LoadResString(9999)
            If Err.Number = 32813 Then
                msg = "Resource not found"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_with_statement() {
        let source = r#"
            With lblStatus
                .Caption = LoadResString(8001)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_select_case() {
        let source = r#"
            Select Case errorType
                Case 1
                    msg = LoadResString(9001)
                Case 2
                    msg = LoadResString(9002)
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_elseif() {
        let source = r#"
            If lang = "en" Then
                msg = LoadResString(10001)
            ElseIf lang = "es" Then
                msg = LoadResString(11001)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_concatenation() {
        let source = r#"
            Dim fullMsg As String
            fullMsg = LoadResString(12001) & vbCrLf & details
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_parentheses() {
        let source = r#"
            msg = (LoadResString(13001))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_iif() {
        let source = r#"
            msg = IIf(success, LoadResString(14001), LoadResString(14002))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_in_class() {
        let source = r#"
            Private Sub Class_Initialize()
                m_errorMsg = LoadResString(15001)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_function_argument() {
        let source = r#"
            Call ShowMessage(LoadResString(16001))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_property_assignment() {
        let source = r#"
            MyObject.Message = LoadResString(17001)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_array_assignment() {
        let source = r#"
            messages(i) = LoadResString(18000 + i)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_while_wend() {
        let source = r#"
            While index < maxStrings
                text = LoadResString(19000 + index)
                index = index + 1
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_do_while() {
        let source = r#"
            Do While hasMore
                currentMsg = LoadResString(GetNextID())
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_do_until() {
        let source = r#"
            Do Until loaded
                On Error Resume Next
                msg = LoadResString(resID)
                loaded = (Err.Number = 0)
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_constants() {
        let source = r#"
            Const MSG_ERROR = 20001
            MsgBox LoadResString(MSG_ERROR), vbCritical
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_addition() {
        let source = r#"
            Dim baseID As Integer
            baseID = 21000
            msg = LoadResString(baseID + offset)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_replace() {
        let source = r#"
            Dim template As String
            template = LoadResString(22001)
            msg = Replace(template, "{0}", userName)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_collection_add() {
        let source = r#"
            messages.Add LoadResString(23001), "WelcomeMsg"
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_debug_print() {
        let source = r#"
            Debug.Print LoadResString(24001)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_loadresstring_tooltip() {
        let source = r#"
            cmdSave.ToolTipText = LoadResString(25001)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResString"));
        assert!(text.contains("Identifier"));
    }
}
