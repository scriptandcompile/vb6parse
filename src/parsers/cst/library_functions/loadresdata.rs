//! # `LoadResData` Function
//!
//! Returns the data stored in a resource (.res) file.
//!
//! ## Syntax
//!
//! ```vb
//! LoadResData(index, format)
//! ```
//!
//! ## Parameters
//!
//! - `index` (Required): `Integer` or `String` identifying the resource
//!   - Can be a numeric ID or string name
//!   - Must match the ID/name used when the resource was compiled
//! - `format` (Required): `Integer` specifying the format of the resource data
//!   - Custom user-defined format number (1-32767)
//!   - Cannot use predefined formats (use `LoadResPicture`/`LoadResString` instead)
//!
//! ## Return Value
//!
//! Returns a Byte array (Variant containing Byte array):
//! - Contains the raw binary data from the resource file
//! - Array is zero-based
//! - Returns Empty if resource not found (in some VB versions)
//! - May raise error 32813 if resource not found
//! - Data is returned exactly as stored in .res file
//!
//! ## Remarks
//!
//! The `LoadResData` function loads custom binary data from resources:
//!
//! - Loads custom data from compiled resource (.res) files
//! - Resource file must be linked to project at compile time
//! - Returns data as Byte array (`Variant`/`Byte()`)
//! - Used for custom binary resources (not bitmaps/icons/strings)
//! - For images, use `LoadResPicture` instead
//! - For strings, use `LoadResString` instead
//! - Resource files created with Resource Editor or RC.EXE
//! - Index can be numeric ID or string name
//! - Format must be custom type (not 1=Cursor, 2=Bitmap, 3=Icon, etc.)
//! - Useful for embedding files (sounds, data, etc.)
//! - Data embedded in compiled EXE (no external files needed)
//! - More secure than external files (can't be easily modified)
//! - Faster access than loading from disk
//! - Resource file added via Project > Add File or in .vbp
//! - Only one resource file per project
//! - Changes to .res file require recompile
//! - Error 32813: "Resource not found" if ID/format don't match
//! - Error 48: "Error loading from file" if resource file corrupt
//!
//! ## Typical Uses
//!
//! 1. **Load Binary Data**
//!    ```vb
//!    Dim data() As Byte
//!    data = LoadResData(101, 256) ' Custom format 256
//!    ```
//!
//! 2. **Load Sound File**
//!    ```vb
//!    Dim wavData() As Byte
//!    wavData = LoadResData("STARTUP_SOUND", 257)
//!    ```
//!
//! 3. **Load Configuration Data**
//!    ```vb
//!    Dim configData() As Byte
//!    configData = LoadResData("CONFIG", 300)
//!    ```
//!
//! 4. **Load Binary Template**
//!    ```vb
//!    Dim template() As Byte
//!    template = LoadResData(1001, 400)
//!    ```
//!
//! 5. **Load Custom File Type**
//!    ```vb
//!    Dim xmlData() As Byte
//!    xmlData = LoadResData("SCHEMA", 500)
//!    ```
//!
//! 6. **Load Multiple Resources**
//!    ```vb
//!    For i = 1 To 5
//!        data = LoadResData(i, 256)
//!        ProcessData data
//!    Next i
//!    ```
//!
//! 7. **Load Named Resource**
//!    ```vb
//!    Dim helpData() As Byte
//!    helpData = LoadResData("HELP_FILE", 600)
//!    ```
//!
//! 8. **Conditional Loading**
//!    ```vb
//!    If useCustomTheme Then
//!        themeData = LoadResData("DARK_THEME", 700)
//!    End If
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Loading and Using Binary Data
//! ```vb
//! ' Load binary data from resource
//! Dim resourceData() As Byte
//! resourceData = LoadResData(101, 256)
//!
//! ' Use the data
//! Dim fileNum As Integer
//! fileNum = FreeFile
//! Open "C:\output.dat" For Binary As #fileNum
//! Put #fileNum, , resourceData
//! Close #fileNum
//! ```
//!
//! ### Example 2: Loading Sound Data
//! ```vb
//! ' Load WAV file from resources
//! Dim wavData() As Byte
//! wavData = LoadResData("BEEP", 257)
//!
//! ' Play using Windows API
//! Call sndPlaySound(wavData(0), SND_ASYNC Or SND_MEMORY)
//! ```
//!
//! ### Example 3: Error Handling
//! ```vb
//! On Error Resume Next
//! Dim data() As Byte
//! data = LoadResData(999, 256)
//! If Err.Number = 32813 Then
//!     MsgBox "Resource not found!", vbCritical
//!     Err.Clear
//! ElseIf Err.Number <> 0 Then
//!     MsgBox "Error loading resource: " & Err.Description, vbCritical
//!     Err.Clear
//! End If
//! ```
//!
//! ### Example 4: Converting to String
//! ```vb
//! ' Load text data stored as binary
//! Dim textData() As Byte
//! textData = LoadResData("README", 300)
//!
//! ' Convert byte array to string
//! Dim content As String
//! content = StrConv(textData, vbUnicode)
//! MsgBox content
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `SafeLoadResData`
//! ```vb
//! Function SafeLoadResData(ByVal resID As Variant, _
//!                          ByVal resFormat As Integer, _
//!                          ByRef outData() As Byte) As Boolean
//!     On Error Resume Next
//!     outData = LoadResData(resID, resFormat)
//!     SafeLoadResData = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 2: `LoadResDataToFile`
//! ```vb
//! Sub LoadResDataToFile(ByVal resID As Variant, _
//!                       ByVal resFormat As Integer, _
//!                       ByVal filename As String)
//!     Dim data() As Byte
//!     Dim fileNum As Integer
//!     
//!     data = LoadResData(resID, resFormat)
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     Put #fileNum, , data
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Pattern 3: `GetResDataSize`
//! ```vb
//! Function GetResDataSize(ByVal resID As Variant, _
//!                         ByVal resFormat As Integer) As Long
//!     On Error Resume Next
//!     Dim data() As Byte
//!     data = LoadResData(resID, resFormat)
//!     
//!     If Err.Number = 0 Then
//!         GetResDataSize = UBound(data) + 1
//!     Else
//!         GetResDataSize = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `LoadResDataAsString`
//! ```vb
//! Function LoadResDataAsString(ByVal resID As Variant, _
//!                              ByVal resFormat As Integer) As String
//!     Dim data() As Byte
//!     data = LoadResData(resID, resFormat)
//!     LoadResDataAsString = StrConv(data, vbUnicode)
//! End Function
//! ```
//!
//! ### Pattern 5: `ResDataExists`
//! ```vb
//! Function ResDataExists(ByVal resID As Variant, _
//!                        ByVal resFormat As Integer) As Boolean
//!     On Error Resume Next
//!     Dim data() As Byte
//!     data = LoadResData(resID, resFormat)
//!     ResDataExists = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 6: `LoadMultipleResources`
//! ```vb
//! Function LoadMultipleResources(startID As Integer, _
//!                                endID As Integer, _
//!                                resFormat As Integer) As Collection
//!     Dim col As New Collection
//!     Dim i As Integer
//!     Dim data() As Byte
//!     
//!     On Error Resume Next
//!     For i = startID To endID
//!         data = LoadResData(i, resFormat)
//!         If Err.Number = 0 Then
//!             col.Add data
//!         End If
//!         Err.Clear
//!     Next i
//!     
//!     Set LoadMultipleResources = col
//! End Function
//! ```
//!
//! ### Pattern 7: `CompareResData`
//! ```vb
//! Function CompareResData(ByVal resID1 As Variant, _
//!                         ByVal resID2 As Variant, _
//!                         ByVal resFormat As Integer) As Boolean
//!     Dim data1() As Byte, data2() As Byte
//!     Dim i As Long
//!     
//!     data1 = LoadResData(resID1, resFormat)
//!     data2 = LoadResData(resID2, resFormat)
//!     
//!     If UBound(data1) <> UBound(data2) Then
//!         CompareResData = False
//!         Exit Function
//!     End If
//!     
//!     For i = 0 To UBound(data1)
//!         If data1(i) <> data2(i) Then
//!             CompareResData = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     CompareResData = True
//! End Function
//! ```
//!
//! ### Pattern 8: `CopyResDataToArray`
//! ```vb
//! Sub CopyResDataToArray(ByVal resID As Variant, _
//!                        ByVal resFormat As Integer, _
//!                        ByRef targetArray() As Byte)
//!     Dim source() As Byte
//!     Dim i As Long
//!     
//!     source = LoadResData(resID, resFormat)
//!     ReDim targetArray(LBound(source) To UBound(source))
//!     
//!     For i = LBound(source) To UBound(source)
//!         targetArray(i) = source(i)
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 9: `LoadResDataWithFallback`
//! ```vb
//! Function LoadResDataWithFallback(ByVal primaryID As Variant, _
//!                                  ByVal fallbackID As Variant, _
//!                                  ByVal resFormat As Integer) As Byte()
//!     On Error Resume Next
//!     
//!     LoadResDataWithFallback = LoadResData(primaryID, resFormat)
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         LoadResDataWithFallback = LoadResData(fallbackID, resFormat)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `CachedResDataLoader`
//! ```vb
//! Dim resCache As New Collection
//!
//! Function LoadResDataCached(ByVal resID As Variant, _
//!                            ByVal resFormat As Integer) As Byte()
//!     Dim key As String
//!     On Error Resume Next
//!     
//!     key = CStr(resID) & "_" & CStr(resFormat)
//!     
//!     ' Try cache first
//!     LoadResDataCached = resCache(key)
//!     If Err.Number <> 0 Then
//!         ' Not in cache, load and cache it
//!         Err.Clear
//!         LoadResDataCached = LoadResData(resID, resFormat)
//!         resCache.Add LoadResDataCached, key
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Resource Manager Class
//! ```vb
//! ' Class: ResourceManager
//! Private m_cache As Collection
//!
//! Private Sub Class_Initialize()
//!     Set m_cache = New Collection
//! End Sub
//!
//! Public Function LoadData(ByVal resID As Variant, _
//!                          ByVal resFormat As Integer) As Byte()
//!     Dim key As String
//!     On Error Resume Next
//!     
//!     key = CStr(resID) & "_" & CStr(resFormat)
//!     LoadData = m_cache(key)
//!     
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         LoadData = LoadResData(resID, resFormat)
//!         If Err.Number = 0 Then
//!             m_cache.Add LoadData, key
//!         Else
//!             Err.Raise vbObjectError + 1000, "ResourceManager", _
//!                       "Failed to load resource"
//!         End If
//!     End If
//! End Function
//!
//! Public Function GetAsString(ByVal resID As Variant, _
//!                             ByVal resFormat As Integer) As String
//!     Dim data() As Byte
//!     data = LoadData(resID, resFormat)
//!     GetAsString = StrConv(data, vbUnicode)
//! End Function
//!
//! Public Sub SaveToFile(ByVal resID As Variant, _
//!                       ByVal resFormat As Integer, _
//!                       ByVal filename As String)
//!     Dim data() As Byte
//!     Dim fileNum As Integer
//!     
//!     data = LoadData(resID, resFormat)
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     Put #fileNum, , data
//!     Close #fileNum
//! End Sub
//!
//! Public Sub ClearCache()
//!     Set m_cache = New Collection
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     Set m_cache = Nothing
//! End Sub
//! ```
//!
//! ### Example 2: Sound Player with Resources
//! ```vb
//! ' Module with API declarations
//! Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" ( _
//!     lpszSoundName As Any, ByVal uFlags As Long) As Long
//!
//! Private Const SND_ASYNC = &H1
//! Private Const SND_MEMORY = &H4
//!
//! ' Sound resource IDs
//! Private Const RES_SOUND_BEEP = 101
//! Private Const RES_SOUND_CLICK = 102
//! Private Const RES_SOUND_ERROR = 103
//! Private Const RES_FORMAT_WAVE = 257
//!
//! Public Sub PlayResourceSound(ByVal soundID As Integer)
//!     Dim soundData() As Byte
//!     On Error Resume Next
//!     
//!     soundData = LoadResData(soundID, RES_FORMAT_WAVE)
//!     If Err.Number = 0 Then
//!         Call sndPlaySound(soundData(0), SND_ASYNC Or SND_MEMORY)
//!     Else
//!         Debug.Print "Sound not found: " & soundID
//!     End If
//! End Sub
//!
//! Public Sub PlayBeep()
//!     PlayResourceSound RES_SOUND_BEEP
//! End Sub
//!
//! Public Sub PlayClick()
//!     PlayResourceSound RES_SOUND_CLICK
//! End Sub
//!
//! Public Sub PlayError()
//!     PlayResourceSound RES_SOUND_ERROR
//! End Sub
//! ```
//!
//! ### Example 3: Configuration Manager
//! ```vb
//! ' Class: ConfigManager
//! Private m_config As String
//!
//! Public Sub LoadConfig()
//!     Dim configData() As Byte
//!     On Error Resume Next
//!     
//!     ' Try to load from resource
//!     configData = LoadResData("CONFIG", 300)
//!     If Err.Number = 0 Then
//!         m_config = StrConv(configData, vbUnicode)
//!     Else
//!         ' Use default config
//!         m_config = GetDefaultConfig()
//!     End If
//!     Err.Clear
//! End Sub
//!
//! Public Function GetSetting(ByVal key As String) As String
//!     Dim lines() As String
//!     Dim i As Long
//!     Dim pos As Long
//!     
//!     lines = Split(m_config, vbCrLf)
//!     For i = 0 To UBound(lines)
//!         If Left(lines(i), Len(key) + 1) = key & "=" Then
//!             GetSetting = Mid(lines(i), Len(key) + 2)
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//!
//! Private Function GetDefaultConfig() As String
//!     GetDefaultConfig = "Version=1.0" & vbCrLf & _
//!                       "Language=English" & vbCrLf & _
//!                       "Theme=Default"
//! End Function
//! ```
//!
//! ### Example 4: Template Engine
//! ```vb
//! ' Class: TemplateEngine
//! Private Const RES_FORMAT_TEMPLATE = 400
//!
//! Public Function ProcessTemplate(ByVal templateID As Variant, _
//!                                  ParamArray values()) As String
//!     Dim template As String
//!     Dim data() As Byte
//!     Dim i As Long
//!     Dim result As String
//!     
//!     ' Load template from resources
//!     data = LoadResData(templateID, RES_FORMAT_TEMPLATE)
//!     template = StrConv(data, vbUnicode)
//!     
//!     result = template
//!     
//!     ' Replace placeholders with values
//!     For i = LBound(values) To UBound(values)
//!         result = Replace(result, "{" & i & "}", CStr(values(i)))
//!     Next i
//!     
//!     ProcessTemplate = result
//! End Function
//!
//! Public Function LoadTemplate(ByVal templateID As Variant) As String
//!     Dim data() As Byte
//!     data = LoadResData(templateID, RES_FORMAT_TEMPLATE)
//!     LoadTemplate = StrConv(data, vbUnicode)
//! End Function
//!
//! Public Function HasTemplate(ByVal templateID As Variant) As Boolean
//!     On Error Resume Next
//!     Dim data() As Byte
//!     data = LoadResData(templateID, RES_FORMAT_TEMPLATE)
//!     HasTemplate = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 32813: Resource not found
//! On Error Resume Next
//! data = LoadResData(999, 256)
//! If Err.Number = 32813 Then
//!     MsgBox "Resource ID 999 not found!"
//! End If
//!
//! ' Error 48: Error loading from file
//! data = LoadResData(101, 256)
//! If Err.Number = 48 Then
//!     MsgBox "Resource file is corrupt or missing!"
//! End If
//!
//! ' Safe loading with error handling
//! Function TryLoadResData(ByVal resID As Variant, _
//!                         ByVal resFormat As Integer, _
//!                         ByRef outData() As Byte) As Boolean
//!     On Error Resume Next
//!     outData = LoadResData(resID, resFormat)
//!     TryLoadResData = (Err.Number = 0)
//!     If Err.Number <> 0 Then
//!         Debug.Print "Error loading resource: " & Err.Description
//!     End If
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Access**: Resources are embedded in EXE (very fast loading)
//! - **Memory Overhead**: Data loaded into memory when accessed
//! - **No Caching**: Each call loads fresh copy (implement caching if needed)
//! - **Compile Time**: Large resources increase EXE size and compile time
//! - **One Resource File**: Only one .res file per project (combine all resources)
//! - **Array Copy**: Returns copy of data (not reference)
//!
//! ## Best Practices
//!
//! 1. **Always handle errors** - resource might not exist
//! 2. **Use constants** for resource IDs and formats
//! 3. **Cache frequently used data** to avoid repeated loading
//! 4. **Use meaningful names** for string-based resource IDs
//! 5. **Document resource IDs** in code or separate file
//! 6. **Keep resources organized** by format number
//! 7. **Test resource loading** during development
//! 8. **Validate data size** after loading if needed
//! 9. **Consider memory usage** for large resources
//! 10. **Use appropriate formats** - `LoadResPicture` for images, `LoadResString` for text
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Return Type | Data Type |
//! |----------|---------|-------------|-----------|
//! | **`LoadResData`** | Load custom binary data | Byte array | Any binary data |
//! | **`LoadResPicture`** | Load image from resources | `StdPicture` | BMP, ICO, CUR |
//! | **`LoadResString`** | Load string from resources | String | Text strings |
//! | **`LoadPicture`** | Load image from file | `StdPicture` | External file |
//!
//! ## `LoadResData` vs `LoadResPicture` vs `LoadResString`
//!
//! ```vb
//! ' LoadResData - custom binary data
//! Dim binData() As Byte
//! binData = LoadResData(101, 256)
//!
//! ' LoadResPicture - images
//! Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!
//! ' LoadResString - strings
//! Dim msg As String
//! msg = LoadResString(101)
//! ```
//!
//! **When to use each:**
//! - **`LoadResData`**: Custom binary files, sound files, configuration data
//! - **`LoadResPicture`**: Images (bitmaps, icons, cursors)
//! - **`LoadResString`**: Localized text strings
//!
//! ## Platform Notes
//!
//! - Available in VB6 (not in early VB versions)
//! - Requires resource file (.res) linked to project
//! - Resource file created with Resource Editor or RC.EXE
//! - Only one resource file per project
//! - Resources embedded in compiled EXE/DLL
//! - Returns Variant containing Byte array
//! - Array is always zero-based (0 to n-1)
//! - Custom formats: 1-32767 (avoid predefined formats)
//! - Maximum resource size limited by available memory
//!
//! ## Limitations
//!
//! - **One resource file**: Only one .res file per project
//! - **Compile time**: Must recompile to update resources
//! - **No modification**: Cannot modify resources at runtime
//! - **Limited tools**: VB6 Resource Editor is basic
//! - **Format restrictions**: Cannot use predefined formats (1-16)
//! - **No compression**: Resources stored uncompressed in EXE
//! - **No encryption**: Resources easily extractable from EXE
//! - **Memory copy**: Always returns copy of data (not reference)
//! - **Error messages**: Limited error information
//! - **No streaming**: Entire resource loaded into memory
//!
//! ## Related Functions
//!
//! - `LoadResPicture`: Load picture from resource file
//! - `LoadResString`: Load string from resource file
//! - `LoadPicture`: Load picture from external file
//! - `StrConv`: Convert byte array to string
//! - `FreeFile`: Get file number for saving resource data

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn loadresdata_basic() {
        let source = r#"
            Dim data() As Byte
            data = LoadResData(101, 256)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_string_id() {
        let source = r#"
            data = LoadResData("SOUND", 257)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_if_statement() {
        let source = r#"
            If hasResource Then
                data = LoadResData(resID, resFormat)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_function_return() {
        let source = r#"
            Function GetResourceData() As Byte()
                GetResourceData = LoadResData(101, 256)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_for_loop() {
        let source = r#"
            For i = 1 To 10
                resData = LoadResData(i, 256)
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_error_handling() {
        let source = r#"
            On Error Resume Next
            data = LoadResData(999, 256)
            If Err.Number = 32813 Then
                MsgBox "Resource not found"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_with_statement() {
        let source = r#"
            With resourceManager
                .data = LoadResData(101, 256)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_array_assignment() {
        let source = r#"
            Dim resources(1 To 5) As Variant
            resources(i) = LoadResData(i, 256)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_select_case() {
        let source = r#"
            Select Case resourceType
                Case 1
                    data = LoadResData(101, 256)
                Case 2
                    data = LoadResData(102, 257)
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_elseif() {
        let source = r#"
            If mode = 1 Then
                data = LoadResData(101, 256)
            ElseIf mode = 2 Then
                data = LoadResData(102, 256)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_strconv() {
        let source = r#"
            Dim textData As String
            textData = StrConv(LoadResData("TEXT", 300), vbUnicode)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_parentheses() {
        let source = r#"
            data = (LoadResData(101, 256))
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_iif() {
        let source = r#"
            data = IIf(useCustom, LoadResData(101, 256), LoadResData(1, 256))
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_in_class() {
        let source = r#"
            Private Sub Class_Initialize()
                m_data = LoadResData("DEFAULT", 256)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_function_argument() {
        let source = r#"
            Call ProcessData(LoadResData(101, 256))
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_property_assignment() {
        let source = r#"
            MyObject.ResourceData = LoadResData(101, 256)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_ubound() {
        let source = r#"
            Dim size As Long
            size = UBound(LoadResData(101, 256)) + 1
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_while_wend() {
        let source = r#"
            While index < maxResources
                data = LoadResData(index, 256)
                index = index + 1
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_do_while() {
        let source = r#"
            Do While hasMore
                currentData = LoadResData(GetNextID(), format)
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_do_until() {
        let source = r#"
            Do Until loaded
                On Error Resume Next
                data = LoadResData(resID, format)
                loaded = (Err.Number = 0)
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_constants() {
        let source = r#"
            Const RES_FORMAT_WAVE = 257
            data = LoadResData(101, RES_FORMAT_WAVE)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_file_write() {
        let source = r#"
            Dim fileNum As Integer
            fileNum = FreeFile
            Put #fileNum, , LoadResData(101, 256)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_concatenation() {
        let source = r#"
            Dim id As String
            id = "RES_" & resNum
            data = LoadResData(id, 256)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_collection_add() {
        let source = r#"
            resources.Add LoadResData(i, 256), "Resource" & i
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_debug_print() {
        let source = r#"
            Debug.Print "Size: " & UBound(LoadResData(101, 256)) + 1
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_comparison() {
        let source = r#"
            If LoadResData(101, 256)(0) = &H4D Then
                MsgBox "Valid header"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadresdata_msgbox() {
        let source = r#"
            MsgBox "Loaded " & UBound(LoadResData(101, 256)) + 1 & " bytes"
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResData"));
        assert!(text.contains("Identifier"));
    }
}
