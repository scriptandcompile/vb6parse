//! # `LoadResPicture` Function
//!
//! Returns a picture object (`StdPicture`) containing an image from a resource (.res) file.
//!
//! ## Syntax
//!
//! ```vb
//! LoadResPicture(index, format)
//! ```
//!
//! ## Parameters
//!
//! - `index` (Required): Integer or String identifying the picture resource
//!   - Can be a numeric ID or string name
//!   - Must match the ID/name used when the resource was compiled
//! - `format` (Required): Integer specifying the format of the picture
//!   - `vbResBitmap` (0): Bitmap (.bmp)
//!   - `vbResIcon` (1): Icon (.ico)
//!   - `vbResCursor` (2): Cursor (.cur)
//!
//! ## Return Value
//!
//! Returns a `StdPicture` object:
//! - Picture object containing the loaded image from resources
//! - Object can be assigned to Picture properties of controls
//! - Object implements `IPicture` interface
//! - Returns Nothing if resource not found (some VB versions)
//! - May raise error 32813 if resource not found
//!
//! ## Remarks
//!
//! The `LoadResPicture` function loads images from embedded resources:
//!
//! - Loads images from compiled resource (.res) files
//! - Resource file must be linked to project at compile time
//! - Supports BMP, ICO, and CUR formats only
//! - Does NOT support JPG, GIF, or PNG
//! - Returns `StdPicture` object implementing `IPicture`
//! - Alternative to `LoadPicture` for embedded images
//! - No external files needed at runtime
//! - Faster than loading from disk
//! - More secure (can't be modified by users)
//! - Resources embedded in compiled EXE/DLL
//! - Only one resource file per project
//! - Resource file added via Project > Add File
//! - Resource files created with Resource Editor or RC.EXE
//! - Index can be numeric ID or string name
//! - Format constants: vbResBitmap, vbResIcon, vbResCursor
//! - Error 32813: "Resource not found" if ID/format don't match
//! - Error 48: "Error loading from file" if resource file corrupt
//! - Pictures are not cached (loaded each time)
//! - Set object = Nothing to release memory
//! - Common in `Form_Load` for initial graphics
//! - Used with Image, `PictureBox`, and Form.Picture
//! - Preferred for distribution (no external image files)
//!
//! ## Typical Uses
//!
//! 1. **Load Bitmap to `PictureBox`**
//!    ```vb
//!    Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!    ```
//!
//! 2. **Load Icon to Image Control**
//!    ```vb
//!    Image1.Picture = LoadResPicture(102, vbResIcon)
//!    ```
//!
//! 3. **Load Form Background**
//!    ```vb
//!    Me.Picture = LoadResPicture("BACKGROUND", vbResBitmap)
//!    ```
//!
//! 4. **Load Cursor**
//!    ```vb
//!    Me.MousePointer = vbCustom
//!    Me.MouseIcon = LoadResPicture(103, vbResCursor)
//!    ```
//!
//! 5. **Load Named Resource**
//!    ```vb
//!    imgLogo.Picture = LoadResPicture("LOGO", vbResBitmap)
//!    ```
//!
//! 6. **Conditional Image Loading**
//!    ```vb
//!    If mode = "dark" Then
//!        Picture1.Picture = LoadResPicture(201, vbResBitmap)
//!    Else
//!        Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!    End If
//!    ```
//!
//! 7. **Button Icons**
//!    ```vb
//!    cmdSave.Picture = LoadResPicture(104, vbResIcon)
//!    ```
//!
//! 8. **Multiple Images in Loop**
//!    ```vb
//!    For i = 1 To 5
//!        imgArray(i).Picture = LoadResPicture(100 + i, vbResBitmap)
//!    Next i
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Picture Loading
//! ```vb
//! ' Load bitmap from resources
//! Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!
//! ' Load icon
//! Image1.Picture = LoadResPicture(102, vbResIcon)
//!
//! ' Load using string name
//! Picture2.Picture = LoadResPicture("SPLASH", vbResBitmap)
//! ```
//!
//! ### Example 2: Form Initialization
//! ```vb
//! Private Sub Form_Load()
//!     ' Load form background
//!     Me.Picture = LoadResPicture(101, vbResBitmap)
//!     
//!     ' Load toolbar icons
//!     cmdNew.Picture = LoadResPicture(201, vbResIcon)
//!     cmdOpen.Picture = LoadResPicture(202, vbResIcon)
//!     cmdSave.Picture = LoadResPicture(203, vbResIcon)
//! End Sub
//! ```
//!
//! ### Example 3: Error Handling
//! ```vb
//! On Error Resume Next
//! Picture1.Picture = LoadResPicture(999, vbResBitmap)
//! If Err.Number = 32813 Then
//!     MsgBox "Resource not found!", vbCritical
//!     Err.Clear
//! ElseIf Err.Number <> 0 Then
//!     MsgBox "Error loading resource: " & Err.Description, vbCritical
//!     Err.Clear
//! End If
//! ```
//!
//! ### Example 4: Dynamic Loading
//! ```vb
//! Dim imageID As Integer
//! imageID = 101 + selectedIndex
//! Picture1.Picture = LoadResPicture(imageID, vbResBitmap)
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `SafeLoadResPicture`
//! ```vb
//! Function SafeLoadResPicture(ByVal resID As Variant, _
//!                             ByVal resFormat As Integer, _
//!                             ByVal ctrl As Object) As Boolean
//!     On Error Resume Next
//!     Set ctrl.Picture = LoadResPicture(resID, resFormat)
//!     SafeLoadResPicture = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 2: `PreloadResourcePictures`
//! ```vb
//! Dim preloadedPics() As StdPicture
//!
//! Sub PreloadResourcePictures()
//!     Dim i As Long
//!     ReDim preloadedPics(1 To 5)
//!     
//!     For i = 1 To 5
//!         Set preloadedPics(i) = LoadResPicture(100 + i, vbResBitmap)
//!     Next i
//! End Sub
//!
//! Sub ShowPreloadedImage(ByVal index As Long)
//!     If index >= 1 And index <= UBound(preloadedPics) Then
//!         Set Picture1.Picture = preloadedPics(index)
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 3: `LoadResPictureWithDefault`
//! ```vb
//! Function LoadResPictureWithDefault(ByVal resID As Variant, _
//!                                    ByVal resFormat As Integer, _
//!                                    ByVal defaultID As Variant) As StdPicture
//!     On Error Resume Next
//!     
//!     Set LoadResPictureWithDefault = LoadResPicture(resID, resFormat)
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         Set LoadResPictureWithDefault = LoadResPicture(defaultID, resFormat)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `LoadResByName`
//! ```vb
//! Function LoadResByName(ByVal resName As String, _
//!                        ByVal resFormat As Integer) As StdPicture
//!     On Error Resume Next
//!     Set LoadResByName = LoadResPicture(resName, resFormat)
//!     
//!     If Err.Number <> 0 Then
//!         Debug.Print "Failed to load resource: " & resName
//!         Set LoadResByName = Nothing
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `ToggleResPicture`
//! ```vb
//! Dim currentState As Boolean
//!
//! Sub ToggleResPicture()
//!     If currentState Then
//!         Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!     Else
//!         Picture1.Picture = LoadResPicture(102, vbResBitmap)
//!     End If
//!     currentState = Not currentState
//! End Sub
//! ```
//!
//! ### Pattern 6: `LoadThemeResources`
//! ```vb
//! Enum ThemeType
//!     tmLight = 0
//!     tmDark = 1
//! End Enum
//!
//! Sub LoadThemeResources(theme As ThemeType)
//!     Dim baseID As Integer
//!     baseID = IIf(theme = tmDark, 200, 100)
//!     
//!     Me.Picture = LoadResPicture(baseID + 1, vbResBitmap)
//!     Picture1.Picture = LoadResPicture(baseID + 2, vbResBitmap)
//!     Picture2.Picture = LoadResPicture(baseID + 3, vbResBitmap)
//! End Sub
//! ```
//!
//! ### Pattern 7: `ResExists`
//! ```vb
//! Function ResExists(ByVal resID As Variant, _
//!                    ByVal resFormat As Integer) As Boolean
//!     On Error Resume Next
//!     Dim pic As StdPicture
//!     Set pic = LoadResPicture(resID, resFormat)
//!     ResExists = (Err.Number = 0)
//!     Set pic = Nothing
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 8: `LoadAllResourceIcons`
//! ```vb
//! Function LoadAllResourceIcons(startID As Integer, _
//!                               endID As Integer) As Collection
//!     Dim col As New Collection
//!     Dim i As Integer
//!     Dim pic As StdPicture
//!     
//!     On Error Resume Next
//!     For i = startID To endID
//!         Set pic = LoadResPicture(i, vbResIcon)
//!         If Err.Number = 0 Then
//!             col.Add pic
//!         End If
//!         Err.Clear
//!     Next i
//!     
//!     Set LoadAllResourceIcons = col
//! End Function
//! ```
//!
//! ### Pattern 9: `SetButtonIcon`
//! ```vb
//! Sub SetButtonIcon(btn As CommandButton, _
//!                   ByVal iconID As Integer, _
//!                   ByVal enabled As Boolean)
//!     On Error Resume Next
//!     
//!     If enabled Then
//!         Set btn.Picture = LoadResPicture(iconID, vbResIcon)
//!     Else
//!         Set btn.Picture = LoadResPicture(iconID + 100, vbResIcon)
//!     End If
//!     
//!     btn.enabled = enabled
//! End Sub
//! ```
//!
//! ### Pattern 10: `LoadResourceArray`
//! ```vb
//! Sub LoadResourceArray(ByRef picArray() As StdPicture, _
//!                       ByVal startID As Integer, _
//!                       ByVal count As Integer, _
//!                       ByVal resFormat As Integer)
//!     Dim i As Integer
//!     ReDim picArray(1 To count)
//!     
//!     On Error Resume Next
//!     For i = 1 To count
//!         Set picArray(i) = LoadResPicture(startID + i - 1, resFormat)
//!         If Err.Number <> 0 Then
//!             Debug.Print "Failed to load resource: " & (startID + i - 1)
//!             Err.Clear
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Resource Picture Manager
//! ```vb
//! ' Class: ResPictureManager
//! Private m_cache As Collection
//!
//! Private Sub Class_Initialize()
//!     Set m_cache = New Collection
//! End Sub
//!
//! Public Function LoadPicture(ByVal resID As Variant, _
//!                             ByVal resFormat As Integer) As StdPicture
//!     Dim key As String
//!     On Error Resume Next
//!     
//!     key = CStr(resID) & "_" & CStr(resFormat)
//!     Set LoadPicture = m_cache(key)
//!     
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         Set LoadPicture = LoadResPicture(resID, resFormat)
//!         If Err.Number = 0 Then
//!             m_cache.Add LoadPicture, key
//!         Else
//!             Err.Raise vbObjectError + 1000, "ResPictureManager", _
//!                       "Failed to load resource"
//!         End If
//!     End If
//! End Function
//!
//! Public Sub AssignToControl(ByVal ctrl As Object, _
//!                            ByVal resID As Variant, _
//!                            ByVal resFormat As Integer)
//!     Set ctrl.Picture = LoadPicture(resID, resFormat)
//! End Sub
//!
//! Public Sub ClearCache()
//!     Dim i As Long
//!     For i = m_cache.Count To 1 Step -1
//!         m_cache.Remove i
//!     Next i
//! End Sub
//!
//! Public Property Get CacheSize() As Long
//!     CacheSize = m_cache.Count
//! End Property
//!
//! Private Sub Class_Terminate()
//!     ClearCache
//!     Set m_cache = Nothing
//! End Sub
//! ```
//!
//! ### Example 2: Image Gallery from Resources
//! ```vb
//! ' Form with Picture1, Timer1, cmdNext, cmdPrev, lblInfo
//! Private Const BASE_IMAGE_ID = 1001
//! Private Const IMAGE_COUNT = 10
//! Private currentIndex As Long
//!
//! Private Sub Form_Load()
//!     currentIndex = 0
//!     ShowCurrentImage
//!     Timer1.Interval = 5000 ' 5 seconds
//!     Timer1.Enabled = True
//! End Sub
//!
//! Private Sub ShowCurrentImage()
//!     Dim imageID As Integer
//!     On Error Resume Next
//!     
//!     imageID = BASE_IMAGE_ID + currentIndex
//!     Set Picture1.Picture = LoadResPicture(imageID, vbResBitmap)
//!     
//!     If Err.Number <> 0 Then
//!         Picture1.Cls
//!         Picture1.Print "Image not found"
//!     Else
//!         lblInfo.Caption = "Image " & (currentIndex + 1) & " of " & IMAGE_COUNT
//!     End If
//!     Err.Clear
//! End Sub
//!
//! Private Sub Timer1_Timer()
//!     NextImage
//! End Sub
//!
//! Private Sub cmdNext_Click()
//!     NextImage
//! End Sub
//!
//! Private Sub cmdPrev_Click()
//!     PrevImage
//! End Sub
//!
//! Private Sub NextImage()
//!     currentIndex = (currentIndex + 1) Mod IMAGE_COUNT
//!     ShowCurrentImage
//! End Sub
//!
//! Private Sub PrevImage()
//!     currentIndex = (currentIndex - 1 + IMAGE_COUNT) Mod IMAGE_COUNT
//!     ShowCurrentImage
//! End Sub
//! ```
//!
//! ### Example 3: Toolbar with Resource Icons
//! ```vb
//! ' Form with toolbar buttons array: cmdTool(0 to 9)
//! Private Type ToolButton
//!     caption As String
//!     iconID As Integer
//!     enabled As Boolean
//!     tooltip As String
//! End Type
//!
//! Private toolConfig() As ToolButton
//!
//! Private Sub Form_Load()
//!     InitializeToolbar
//!     ApplyToolbarConfig
//! End Sub
//!
//! Private Sub InitializeToolbar()
//!     ReDim toolConfig(0 To 9)
//!     
//!     With toolConfig(0)
//!         .caption = "New"
//!         .iconID = 201
//!         .enabled = True
//!         .tooltip = "Create new document"
//!     End With
//!     
//!     With toolConfig(1)
//!         .caption = "Open"
//!         .iconID = 202
//!         .enabled = True
//!         .tooltip = "Open existing document"
//!     End With
//!     
//!     With toolConfig(2)
//!         .caption = "Save"
//!         .iconID = 203
//!         .enabled = False
//!         .tooltip = "Save current document"
//!     End With
//!     
//!     ' ... more buttons
//! End Sub
//!
//! Private Sub ApplyToolbarConfig()
//!     Dim i As Long
//!     On Error Resume Next
//!     
//!     For i = 0 To UBound(toolConfig)
//!         With cmdTool(i)
//!             .caption = toolConfig(i).caption
//!             .enabled = toolConfig(i).enabled
//!             .ToolTipText = toolConfig(i).tooltip
//!             
//!             Set .Picture = LoadResPicture(toolConfig(i).iconID, vbResIcon)
//!             If Err.Number <> 0 Then
//!                 Debug.Print "Failed to load icon: " & toolConfig(i).iconID
//!                 Err.Clear
//!             End If
//!         End With
//!     Next i
//! End Sub
//!
//! Public Sub EnableTool(ByVal index As Long)
//!     If index >= 0 And index <= UBound(toolConfig) Then
//!         toolConfig(index).enabled = True
//!         cmdTool(index).enabled = True
//!     End If
//! End Sub
//!
//! Public Sub DisableTool(ByVal index As Long)
//!     If index >= 0 And index <= UBound(toolConfig) Then
//!         toolConfig(index).enabled = False
//!         cmdTool(index).enabled = False
//!     End If
//! End Sub
//! ```
//!
//! ### Example 4: Multi-State Indicator
//! ```vb
//! ' Form with imgStatus (Image control)
//! Public Enum StatusState
//!     stIdle = 0
//!     stProcessing = 1
//!     stSuccess = 2
//!     stWarning = 3
//!     stError = 4
//! End Enum
//!
//! Private Const RES_STATUS_IDLE = 301
//! Private Const RES_STATUS_PROCESSING = 302
//! Private Const RES_STATUS_SUCCESS = 303
//! Private Const RES_STATUS_WARNING = 304
//! Private Const RES_STATUS_ERROR = 305
//!
//! Private statusIcons() As StdPicture
//! Private currentStatus As StatusState
//!
//! Private Sub Form_Load()
//!     LoadStatusIcons
//!     SetStatus stIdle
//! End Sub
//!
//! Private Sub LoadStatusIcons()
//!     ReDim statusIcons(0 To 4)
//!     
//!     On Error Resume Next
//!     Set statusIcons(stIdle) = LoadResPicture(RES_STATUS_IDLE, vbResIcon)
//!     Set statusIcons(stProcessing) = LoadResPicture(RES_STATUS_PROCESSING, vbResIcon)
//!     Set statusIcons(stSuccess) = LoadResPicture(RES_STATUS_SUCCESS, vbResIcon)
//!     Set statusIcons(stWarning) = LoadResPicture(RES_STATUS_WARNING, vbResIcon)
//!     Set statusIcons(stError) = LoadResPicture(RES_STATUS_ERROR, vbResIcon)
//!     
//!     If Err.Number <> 0 Then
//!         MsgBox "Warning: Some status icons could not be loaded", vbExclamation
//!         Err.Clear
//!     End If
//! End Sub
//!
//! Public Sub SetStatus(ByVal newStatus As StatusState)
//!     currentStatus = newStatus
//!     
//!     If newStatus >= 0 And newStatus <= UBound(statusIcons) Then
//!         If Not statusIcons(newStatus) Is Nothing Then
//!             Set imgStatus.Picture = statusIcons(newStatus)
//!         End If
//!     End If
//! End Sub
//!
//! Public Function GetStatus() As StatusState
//!     GetStatus = currentStatus
//! End Function
//!
//! Private Sub Form_Unload(Cancel As Integer)
//!     Dim i As Long
//!     For i = 0 To UBound(statusIcons)
//!         Set statusIcons(i) = Nothing
//!     Next i
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 32813: Resource not found
//! On Error Resume Next
//! Set pic = LoadResPicture(999, vbResBitmap)
//! If Err.Number = 32813 Then
//!     MsgBox "Resource not found!"
//! End If
//!
//! ' Error 48: Error loading from file
//! Set pic = LoadResPicture(101, vbResBitmap)
//! If Err.Number = 48 Then
//!     MsgBox "Resource file is corrupt or missing!"
//! End If
//!
//! ' Safe loading pattern
//! Function TryLoadResPicture(ByVal resID As Variant, _
//!                            ByVal resFormat As Integer, _
//!                            ByRef pic As StdPicture) As Boolean
//!     On Error Resume Next
//!     Set pic = LoadResPicture(resID, resFormat)
//!     TryLoadResPicture = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Loading**: Resources embedded in EXE (very fast access)
//! - **No File I/O**: No disk access required
//! - **Memory Usage**: Pictures consume memory until released
//! - **No Caching**: Each call loads fresh copy (implement caching if needed)
//! - **Preloading**: Load frequently used images once at startup
//! - **EXE Size**: Large images increase executable size
//!
//! ## Best Practices
//!
//! 1. **Always handle errors** - resource might not exist
//! 2. **Use constants** for resource IDs for maintainability
//! 3. **Preload frequently used images** for better performance
//! 4. **Release memory** by setting picture objects to Nothing when done
//! 5. **Use meaningful names** for string-based resource IDs
//! 6. **Test all resources** during development
//! 7. **Document resource IDs** in code or separate file
//! 8. **Use Resource Editor** to manage resources efficiently
//! 9. **Consider image size** - large bitmaps increase EXE size
//! 10. **Cache in Collection** for images used multiple times
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Source | External Files |
//! |----------|---------|--------|----------------|
//! | **`LoadResPicture`** | Load from resources | Embedded .res | No |
//! | **`LoadPicture`** | Load from file | External file | Yes |
//! | **`LoadResData`** | Load binary data | Embedded .res | No |
//! | **`LoadResString`** | Load string | Embedded .res | No |
//!
//! ## `LoadResPicture` vs `LoadPicture`
//!
//! ```vb
//! ' LoadResPicture - from embedded resources
//! Picture1.Picture = LoadResPicture(101, vbResBitmap)
//!
//! ' LoadPicture - from external file
//! Picture1.Picture = LoadPicture("C:\Images\logo.bmp")
//! ```
//!
//! **When to use each:**
//! - **`LoadResPicture`**: Distribution (no external files), static images, faster loading
//! - **`LoadPicture`**: Dynamic images, user-selected files, easier updates
//!
//! ## Resource Format Constants
//!
//! ```vb
//! ' Format constants
//! Const vbResBitmap = 0  ' Bitmap (.bmp)
//! Const vbResIcon = 1    ' Icon (.ico)
//! Const vbResCursor = 2  ' Cursor (.cur)
//!
//! ' Usage
//! Picture1.Picture = LoadResPicture(101, vbResBitmap)
//! Image1.Picture = LoadResPicture(102, vbResIcon)
//! Me.MouseIcon = LoadResPicture(103, vbResCursor)
//! ```
//!
//! ## Platform Notes
//!
//! - Available in VB6 (not in early VB versions)
//! - Requires resource file (.res) linked to project
//! - Resource file created with Resource Editor or RC.EXE
//! - Only one resource file per project
//! - Resources embedded in compiled EXE/DLL
//! - Returns `StdPicture` object (OLE automation object)
//! - Supports BMP, ICO, CUR formats only
//! - No native support for JPG, GIF, PNG
//! - Format parameter: 0=Bitmap, 1=Icon, 2=Cursor
//! - Icons can contain multiple sizes
//!
//! ## Limitations
//!
//! - **Format Support**: Only BMP, ICO, CUR (no JPG/GIF/PNG)
//! - **One Resource File**: Only one .res file per project
//! - **Compile Time**: Must recompile to update resources
//! - **No Modification**: Cannot modify resources at runtime
//! - **No Caching**: Each call reloads from resource
//! - **EXE Size**: Large images significantly increase EXE size
//! - **No Compression**: Images stored uncompressed
//! - **Limited Editor**: VB6 Resource Editor is basic
//! - **No Metadata**: Cannot read image dimensions before loading
//! - **Memory Usage**: Large images consume significant memory
//!
//! ## Related Functions
//!
//! - `LoadPicture`: Load picture from external file
//! - `LoadResData`: Load custom binary data from resources
//! - `LoadResString`: Load string from resources
//! - `SavePicture`: Save picture object to file
//! - `Set`: Assign object references

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn loadrespicture_basic() {
        let source = r#"
            Set Picture1.Picture = LoadResPicture(101, vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_icon() {
        let source = r#"
            Image1.Picture = LoadResPicture(102, vbResIcon)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_if_statement() {
        let source = r#"
            If hasResource Then
                Picture1.Picture = LoadResPicture(resID, vbResBitmap)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_string_name() {
        let source = r#"
            Picture1.Picture = LoadResPicture("LOGO", vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_form_load() {
        let source = r#"
            Private Sub Form_Load()
                Me.Picture = LoadResPicture(101, vbResBitmap)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_array_assignment() {
        let source = r#"
            Set images(i) = LoadResPicture(100 + i, vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_for_loop() {
        let source = r#"
            For i = 1 To 5
                Set imgArray(i).Picture = LoadResPicture(100 + i, vbResBitmap)
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_function_return() {
        let source = r#"
            Function GetResPicture() As StdPicture
                Set GetResPicture = LoadResPicture(101, vbResBitmap)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_error_handling() {
        let source = r#"
            On Error Resume Next
            Picture1.Picture = LoadResPicture(999, vbResBitmap)
            If Err.Number = 32813 Then
                MsgBox "Resource not found"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_with_statement() {
        let source = r#"
            With Picture1
                .Picture = LoadResPicture(101, vbResBitmap)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_select_case() {
        let source = r#"
            Select Case imageType
                Case 1
                    Picture1.Picture = LoadResPicture(101, vbResBitmap)
                Case 2
                    Picture1.Picture = LoadResPicture(102, vbResIcon)
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_elseif() {
        let source = r#"
            If mode = "dark" Then
                Picture1.Picture = LoadResPicture(201, vbResBitmap)
            ElseIf mode = "light" Then
                Picture1.Picture = LoadResPicture(101, vbResBitmap)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_parentheses() {
        let source = r#"
            Set pic = (LoadResPicture(101, vbResBitmap))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_iif() {
        let source = r#"
            Picture1.Picture = IIf(enabled, LoadResPicture(101, vbResIcon), LoadResPicture(102, vbResIcon))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_in_class() {
        let source = r#"
            Private Sub Class_Initialize()
                Set m_defaultPic = LoadResPicture(101, vbResBitmap)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_function_argument() {
        let source = r#"
            Call SetPicture(LoadResPicture(101, vbResBitmap))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_property_assignment() {
        let source = r#"
            Set MyForm.Picture = LoadResPicture(101, vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_is_nothing() {
        let source = r#"
            Set pic = LoadResPicture(101, vbResBitmap)
            If pic Is Nothing Then
                MsgBox "Failed to load"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_while_wend() {
        let source = r#"
            While index < maxImages
                Set images(index) = LoadResPicture(100 + index, vbResBitmap)
                index = index + 1
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_do_while() {
        let source = r#"
            Do While hasMore
                Set currentPic = LoadResPicture(GetNextID(), vbResBitmap)
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_do_until() {
        let source = r#"
            Do Until loaded
                On Error Resume Next
                Set pic = LoadResPicture(resID, vbResBitmap)
                loaded = (Err.Number = 0)
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_cursor() {
        let source = r#"
            Me.MousePointer = vbCustom
            Me.MouseIcon = LoadResPicture(103, vbResCursor)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_constants() {
        let source = r#"
            Const RES_LOGO = 101
            Picture1.Picture = LoadResPicture(RES_LOGO, vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_addition() {
        let source = r#"
            Dim imageID As Integer
            imageID = 100 + selectedIndex
            Picture1.Picture = LoadResPicture(imageID, vbResBitmap)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_toolbar_button() {
        let source = r#"
            cmdSave.Picture = LoadResPicture(203, vbResIcon)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_debug_print() {
        let source = r#"
            Set pic = LoadResPicture(101, vbResBitmap)
            Debug.Print "Loaded resource 101"
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loadrespicture_msgbox() {
        let source = r#"
            On Error Resume Next
            Set pic = LoadResPicture(101, vbResBitmap)
            If Err.Number <> 0 Then
                MsgBox "Failed to load resource"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LoadResPicture"));
        assert!(text.contains("Identifier"));
    }
}
