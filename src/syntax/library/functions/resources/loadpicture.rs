//! # `LoadPicture` Function
//!
//! Returns a picture object (`StdPicture`) containing an image from a file or memory.
//!
//! ## Syntax
//!
//! ```vb
//! LoadPicture([filename] [, size] [, colordepth] [, x, y])
//! ```
//!
//! ## Parameters
//!
//! - `filename` (Optional): `String` expression specifying the name of the file to load
//!   - Can be bitmap (.bmp), icon (.ico), cursor (.cur), metafile (.wmf), or enhanced metafile (.emf)
//!   - Can be empty string ("") to unload picture (returns Nothing)
//!   - If omitted, returns empty picture object
//! - `size` (Optional): `Long` specifying desired icon/cursor size (only for .ico and .cur files)
//!   - Specified in HIMETRIC units (1 HIMETRIC = 0.01 mm)
//! - `colordepth` (Optional): `Long` specifying desired color depth
//! - `x, y` (Optional): `Long` values specifying desired width and height
//!
//! ## Return Value
//!
//! Returns a `StdPicture` object:
//! - For valid filename: `Picture` object containing the loaded image
//! - For empty string (""): Nothing (unloads picture)
//! - For no parameters: Empty picture object
//! - Object can be assigned to `Picture` properties of controls
//! - Object implements `IPicture` interface
//!
//! ## Remarks
//!
//! The `LoadPicture` function loads graphics from files:
//!
//! - Primary method for loading images in VB6
//! - Supports BMP, ICO, CUR, WMF, and EMF formats
//! - Does NOT support JPG, GIF, or PNG natively
//! - Returns `StdPicture` object implementing `IPicture`
//! - ```LoadPicture("")``` explicitly unloads picture (returns `Nothing`)
//! - ```LoadPicture()``` with no arguments returns empty picture
//! - File must exist or runtime error 53 occurs
//! - Invalid image format causes runtime error 481
//! - Pictures are not cached (loaded each time)
//! - For .ico and .cur files, can specify size/colordepth
//! - Size and colordepth parameters select best match
//! - Pictures consume memory until released
//! - Set object = Nothing to release memory
//! - Common in `Form_Load` for initial graphics
//! - Used with `Image`, `PictureBox`, and `Form.Picture`
//! - Can load from resource file with `LoadResPicture` instead
//! - `SavePicture` is the inverse function (saves to file)
//! - `Picture` objects are `OLE` objects
//!
//! ## Typical Uses
//!
//! 1. **Load `Image` to `PictureBox`**
//!    ```vb
//!    Picture1.Picture = LoadPicture("C:\Images\logo.bmp")
//!    ```
//!
//! 2. **Load `Image` to `Form` Background**
//!    ```vb
//!    Me.Picture = LoadPicture(App.Path & "\background.bmp")
//!    ```
//!
//! 3. **Load `Icon` to `Image` Control**
//!    ```vb
//!    Image1.Picture = LoadPicture("C:\Icons\app.ico")
//!    ```
//!
//! 4. **Clear/Unload `Picture`**
//!    ```vb
//!    Picture1.Picture = LoadPicture("")
//!    ```
//!
//! 5. **Dynamic `Image` Loading**
//!    ```vb
//!    Picture1.Picture = LoadPicture(imageFiles(index))
//!    ```
//!
//! 6. **Conditional `Image` Loading**
//!    ```vb
//!    If fileExists Then
//!        imgStatus.Picture = LoadPicture("ok.bmp")
//!    Else
//!        imgStatus.Picture = LoadPicture("error.bmp")
//!    End If
//!    ```
//!
//! 7. **Button `Icons`**
//!    ```vb
//!    cmdSave.Picture = LoadPicture("save.ico")
//!    ```
//!
//! 8. **Animation Frame Loading**
//!    ```vb
//!    For i = 1 To 10
//!        frames(i) = LoadPicture("frame" & i & ".bmp")
//!    Next i
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Picture Loading
//! ```vb
//! ' Load a bitmap to a picture box
//! Picture1.Picture = LoadPicture("C:\MyApp\logo.bmp")
//!
//! ' Load using relative path
//! Picture1.Picture = LoadPicture(App.Path & "\images\banner.bmp")
//!
//! ' Load icon to image control
//! Image1.Picture = LoadPicture(App.Path & "\icon.ico")
//! ```
//!
//! ### Example 2: Clearing Pictures
//! ```vb
//! ' Unload picture to free memory
//! Picture1.Picture = LoadPicture("")
//!
//! ' Alternative way to clear
//! Set Picture1.Picture = LoadPicture()
//!
//! ' Clear form background
//! Me.Picture = LoadPicture("")
//! ```
//!
//! ### Example 3: Error Handling
//! ```vb
//! On Error Resume Next
//! Picture1.Picture = LoadPicture(filename)
//! If Err.Number <> 0 Then
//!     MsgBox "Could not load image: " & filename, vbCritical
//!     Picture1.Picture = LoadPicture("") ' Clear on error
//!     Err.Clear
//! End If
//! ```
//!
//! ### Example 4: File Existence Check
//! ```vb
//! Dim picPath As String
//! picPath = App.Path & "\logo.bmp"
//!
//! If Dir(picPath) <> "" Then
//!     Picture1.Picture = LoadPicture(picPath)
//! Else
//!     MsgBox "Image file not found!", vbExclamation
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `SafeLoadPicture`
//! ```vb
//! Function SafeLoadPicture(ByVal filename As String, _
//!                          ByVal ctrl As Object) As Boolean
//!     On Error Resume Next
//!     Set ctrl.Picture = LoadPicture(filename)
//!     SafeLoadPicture = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 2: `ImageSwitcher`
//! ```vb
//! Sub SwitchImage(ByVal imageIndex As Long)
//!     Dim imageFiles As Variant
//!     imageFiles = Array("image1.bmp", "image2.bmp", "image3.bmp")
//!     
//!     If imageIndex >= 0 And imageIndex <= UBound(imageFiles) Then
//!         Picture1.Picture = LoadPicture(App.Path & "\" & imageFiles(imageIndex))
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 3: `PreloadImages`
//! ```vb
//! Dim preloadedPics() As StdPicture
//!
//! Sub PreloadImages()
//!     Dim i As Long
//!     ReDim preloadedPics(1 To 5)
//!     
//!     For i = 1 To 5
//!         Set preloadedPics(i) = LoadPicture("pic" & i & ".bmp")
//!     Next i
//! End Sub
//!
//! Sub ShowImage(ByVal index As Long)
//!     Set Picture1.Picture = preloadedPics(index)
//! End Sub
//! ```
//!
//! ### Pattern 4: `TogglePicture`
//! ```vb
//! Dim currentState As Boolean
//!
//! Sub TogglePicture()
//!     If currentState Then
//!         Picture1.Picture = LoadPicture("on.bmp")
//!     Else
//!         Picture1.Picture = LoadPicture("off.bmp")
//!     End If
//!     currentState = Not currentState
//! End Sub
//! ```
//!
//! ### Pattern 5: `LoadPictureWithDefault`
//! ```vb
//! Function LoadPictureWithDefault(ByVal filename As String, _
//!                                  ByVal defaultFile As String) As StdPicture
//!     On Error Resume Next
//!     Set LoadPictureWithDefault = LoadPicture(filename)
//!     If Err.Number <> 0 Then
//!         Set LoadPictureWithDefault = LoadPicture(defaultFile)
//!     End If
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 6: `ClearAllPictures`
//! ```vb
//! Sub ClearAllPictures(frm As Form)
//!     Dim ctrl As Control
//!     
//!     For Each ctrl In frm.Controls
//!         If TypeOf ctrl Is PictureBox Or TypeOf ctrl Is Image Then
//!             ctrl.Picture = LoadPicture("")
//!         End If
//!     Next ctrl
//!     
//!     frm.Picture = LoadPicture("")
//! End Sub
//! ```
//!
//! ### Pattern 7: `LoadFromResourceOrFile`
//! ```vb
//! Function LoadFromResourceOrFile(ByVal resID As Long, _
//!                                  ByVal filename As String) As StdPicture
//!     On Error Resume Next
//!     
//!     ' Try resource first
//!     Set LoadFromResourceOrFile = LoadResPicture(resID, vbResBitmap)
//!     
//!     ' Fall back to file
//!     If Err.Number <> 0 Then
//!         Set LoadFromResourceOrFile = LoadPicture(filename)
//!     End If
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Pattern 8: `ButtonImageState`
//! ```vb
//! Sub SetButtonState(btn As CommandButton, enabled As Boolean)
//!     If enabled Then
//!         btn.Picture = LoadPicture(App.Path & "\btn_enabled.ico")
//!         btn.Enabled = True
//!     Else
//!         btn.Picture = LoadPicture(App.Path & "\btn_disabled.ico")
//!         btn.Enabled = False
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 9: `LoadPictureIfExists`
//! ```vb
//! Function LoadPictureIfExists(ByVal filename As String) As StdPicture
//!     If Dir(filename) <> "" Then
//!         Set LoadPictureIfExists = LoadPicture(filename)
//!     Else
//!         Set LoadPictureIfExists = LoadPicture() ' Empty picture
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `CachedPictureLoader`
//! ```vb
//! Dim pictureCache As New Collection
//!
//! Function LoadPictureCached(ByVal filename As String) As StdPicture
//!     On Error Resume Next
//!     
//!     ' Try to get from cache
//!     Set LoadPictureCached = pictureCache(filename)
//!     
//!     ' If not in cache, load and cache it
//!     If Err.Number <> 0 Then
//!         Set LoadPictureCached = LoadPicture(filename)
//!         pictureCache.Add LoadPictureCached, filename
//!     End If
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: `PictureManager` Class
//! ```vb
//! ' Class: PictureManager
//! Private m_pictures As Collection
//! Private m_basePath As String
//!
//! Private Sub Class_Initialize()
//!     Set m_pictures = New Collection
//!     m_basePath = App.Path & "\images\"
//! End Sub
//!
//! Public Sub LoadPicture(ByVal name As String, ByVal filename As String)
//!     Dim pic As StdPicture
//!     On Error Resume Next
//!     
//!     Set pic = VBA.LoadPicture(m_basePath & filename)
//!     If Err.Number = 0 Then
//!         m_pictures.Add pic, name
//!     Else
//!         Err.Raise vbObjectError + 1000, "PictureManager", _
//!                   "Failed to load: " & filename
//!     End If
//! End Sub
//!
//! Public Function GetPicture(ByVal name As String) As StdPicture
//!     On Error Resume Next
//!     Set GetPicture = m_pictures(name)
//!     If Err.Number <> 0 Then
//!         Err.Raise vbObjectError + 1001, "PictureManager", _
//!                   "Picture not found: " & name
//!     End If
//! End Function
//!
//! Public Sub ClearAll()
//!     Dim i As Long
//!     For i = m_pictures.Count To 1 Step -1
//!         m_pictures.Remove i
//!     Next i
//! End Sub
//!
//! Public Property Get Count() As Long
//!     Count = m_pictures.Count
//! End Property
//!
//! Private Sub Class_Terminate()
//!     ClearAll
//!     Set m_pictures = Nothing
//! End Sub
//! ```
//!
//! ### Example 2: Image Slideshow
//! ```vb
//! ' Form with Picture1, Timer1, cmdNext, cmdPrev
//! Dim imageFiles() As String
//! Dim currentIndex As Long
//!
//! Private Sub Form_Load()
//!     LoadImageList
//!     currentIndex = 0
//!     ShowCurrentImage
//!     Timer1.Interval = 3000 ' 3 seconds
//!     Timer1.Enabled = True
//! End Sub
//!
//! Private Sub LoadImageList()
//!     imageFiles = Array( _
//!         "slide1.bmp", _
//!         "slide2.bmp", _
//!         "slide3.bmp", _
//!         "slide4.bmp", _
//!         "slide5.bmp" _
//!     )
//! End Sub
//!
//! Private Sub ShowCurrentImage()
//!     On Error Resume Next
//!     Picture1.Picture = LoadPicture(App.Path & "\slides\" & imageFiles(currentIndex))
//!     If Err.Number <> 0 Then
//!         Picture1.Cls
//!         Picture1.Print "Image not found"
//!     End If
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
//!     currentIndex = currentIndex + 1
//!     If currentIndex > UBound(imageFiles) Then
//!         currentIndex = 0
//!     End If
//!     ShowCurrentImage
//! End Sub
//!
//! Private Sub PrevImage()
//!     currentIndex = currentIndex - 1
//!     If currentIndex < 0 Then
//!         currentIndex = UBound(imageFiles)
//!     End If
//!     ShowCurrentImage
//! End Sub
//! ```
//!
//! ### Example 3: Dynamic Button Icons
//! ```vb
//! ' Form with command buttons array: cmdAction(0 to 4)
//! Private Type ButtonConfig
//!     caption As String
//!     iconFile As String
//!     enabled As Boolean
//! End Type
//!
//! Private buttonConfigs() As ButtonConfig
//!
//! Private Sub Form_Load()
//!     InitializeButtons
//!     ApplyButtonConfigs
//! End Sub
//!
//! Private Sub InitializeButtons()
//!     ReDim buttonConfigs(0 To 4)
//!     
//!     With buttonConfigs(0)
//!         .caption = "New"
//!         .iconFile = "new.ico"
//!         .enabled = True
//!     End With
//!     
//!     With buttonConfigs(1)
//!         .caption = "Open"
//!         .iconFile = "open.ico"
//!         .enabled = True
//!     End With
//!     
//!     With buttonConfigs(2)
//!         .caption = "Save"
//!         .iconFile = "save.ico"
//!         .enabled = False
//!     End With
//!     
//!     With buttonConfigs(3)
//!         .caption = "Print"
//!         .iconFile = "print.ico"
//!         .enabled = False
//!     End With
//!     
//!     With buttonConfigs(4)
//!         .caption = "Exit"
//!         .iconFile = "exit.ico"
//!         .enabled = True
//!     End With
//! End Sub
//!
//! Private Sub ApplyButtonConfigs()
//!     Dim i As Long
//!     Dim iconPath As String
//!     
//!     For i = 0 To UBound(buttonConfigs)
//!         With cmdAction(i)
//!             .caption = buttonConfigs(i).caption
//!             .enabled = buttonConfigs(i).enabled
//!             
//!             iconPath = App.Path & "\icons\" & buttonConfigs(i).iconFile
//!             If Dir(iconPath) <> "" Then
//!                 .Picture = LoadPicture(iconPath)
//!             End If
//!         End With
//!     Next i
//! End Sub
//!
//! Public Sub SetButtonEnabled(ByVal index As Long, ByVal enabled As Boolean)
//!     If index >= 0 And index <= UBound(buttonConfigs) Then
//!         buttonConfigs(index).enabled = enabled
//!         cmdAction(index).enabled = enabled
//!     End If
//! End Sub
//! ```
//!
//! ### Example 4: Status Indicator
//! ```vb
//! ' Form with imgStatus (Image control)
//! Public Enum StatusType
//!     stIdle = 0
//!     stProcessing = 1
//!     stSuccess = 2
//!     stWarning = 3
//!     stError = 4
//! End Enum
//!
//! Private statusIcons() As StdPicture
//! Private currentStatus As StatusType
//!
//! Private Sub Form_Load()
//!     LoadStatusIcons
//!     SetStatus stIdle
//! End Sub
//!
//! Private Sub LoadStatusIcons()
//!     Dim basePath As String
//!     basePath = App.Path & "\icons\"
//!     
//!     ReDim statusIcons(0 To 4)
//!     
//!     On Error Resume Next
//!     Set statusIcons(stIdle) = LoadPicture(basePath & "idle.ico")
//!     Set statusIcons(stProcessing) = LoadPicture(basePath & "processing.ico")
//!     Set statusIcons(stSuccess) = LoadPicture(basePath & "success.ico")
//!     Set statusIcons(stWarning) = LoadPicture(basePath & "warning.ico")
//!     Set statusIcons(stError) = LoadPicture(basePath & "error.ico")
//!     
//!     If Err.Number <> 0 Then
//!         MsgBox "Warning: Some status icons could not be loaded", vbExclamation
//!         Err.Clear
//!     End If
//! End Sub
//!
//! Public Sub SetStatus(ByVal newStatus As StatusType)
//!     currentStatus = newStatus
//!     
//!     If newStatus >= 0 And newStatus <= UBound(statusIcons) Then
//!         If Not statusIcons(newStatus) Is Nothing Then
//!             Set imgStatus.Picture = statusIcons(newStatus)
//!         End If
//!     End If
//! End Sub
//!
//! Public Function GetStatus() As StatusType
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
//! ' Error 53: File not found
//! On Error Resume Next
//! Picture1.Picture = LoadPicture("nonexistent.bmp")
//! If Err.Number = 53 Then
//!     MsgBox "File not found!"
//! End If
//!
//! ' Error 481: Invalid picture
//! Picture1.Picture = LoadPicture("corrupt.bmp")
//! If Err.Number = 481 Then
//!     MsgBox "Invalid or corrupt image file!"
//! End If
//!
//! ' Error 7: Out of memory (very large images)
//! Picture1.Picture = LoadPicture("huge.bmp")
//! If Err.Number = 7 Then
//!     MsgBox "Insufficient memory to load image!"
//! End If
//!
//! ' Safe loading pattern
//! Function TryLoadPicture(ByVal filename As String, _
//!                         ByRef pic As StdPicture) As Boolean
//!     On Error Resume Next
//!     Set pic = LoadPicture(filename)
//!     TryLoadPicture = (Err.Number = 0)
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **File I/O Overhead**: `LoadPicture` reads from disk (relatively slow)
//! - **Memory Usage**: Large images consume significant memory
//! - **No Caching**: Each call loads from disk (not cached automatically)
//! - **Preloading**: For frequently used images, load once and cache
//! - **Release Memory**: Set ```picture = Nothing``` when done
//! - **Format Matters**: BMP files are larger but faster to load than compressed formats
//! - **Network Paths**: Loading from network shares is much slower
//!
//! ## Best Practices
//!
//! 1. **Always handle errors** when loading pictures (file may not exist)
//! 2. **Check file existence** before loading (use Dir function)
//! 3. **Use relative paths** (`App.Path`) for portability
//! 4. **Release memory** by setting picture objects to `Nothing` when done
//! 5. **Preload frequently used images** for better performance
//! 6. **Clear pictures** with ```LoadPicture("")``` not by setting to `Nothing` directly
//! 7. **Validate file extensions** before attempting to load
//! 8. **Use resource files** (`LoadResPicture`) for embedded images
//! 9. **Handle Out of Memory** errors for large images
//! 10. **Consider image size** - large BMPs consume lots of memory
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Return Type | Notes |
//! |----------|---------|-------------|-------|
//! | **`LoadPicture`** | Load from file | `StdPicture` | Supports BMP, ICO, CUR, WMF, EMF |
//! | **`LoadResPicture`** | Load from resources | `StdPicture` | Embedded in compiled EXE |
//! | **`SavePicture`** | Save to file | N/A (Sub) | Inverse of `LoadPicture` |
//! | **`Set` statement** | Assign picture | N/A | Used to assign picture objects |
//!
//! ## `LoadPicture` vs `LoadResPicture`
//!
//! ```vb
//! ' LoadPicture - from file (requires external file)
//! Picture1.Picture = LoadPicture("logo.bmp")
//!
//! ' LoadResPicture - from resources (embedded in EXE)
//! Picture1.Picture = LoadResPicture(101, vbResBitmap)
//! ```
//!
//! **When to use each:**
//! - **`LoadPicture`**: Images that change, user-selected files, development
//! - **`LoadResPicture`**: Static images, distribution (no external files), resources
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Returns `StdPicture` object (OLE automation object)
//! - Implemented in MSVBVM60.DLL runtime
//! - Supports Windows native image formats
//! - No built-in support for JPG, GIF, PNG (requires third-party controls or APIs)
//! - Icon/Cursor files can contain multiple sizes - `LoadPicture` selects best match
//! - Metafiles (WMF/EMF) are vector formats (scalable)
//! - Bitmaps (BMP) are raster formats (not scalable)
//!
//! ## Limitations
//!
//! - **No JPG/GIF/PNG**: Native support limited to BMP, ICO, CUR, WMF, EMF
//! - **No compression**: BMP files can be very large
//! - **No caching**: Each call reloads from disk
//! - **Memory intensive**: Large images consume lots of RAM
//! - **Synchronous**: Blocks while loading (no async loading)
//! - **No progress**: Cannot monitor loading progress
//! - **No metadata**: Cannot read EXIF or other metadata
//! - **No transformations**: Cannot resize/rotate during load
//! - **File path only**: Cannot load from memory or stream
//! - **Error codes limited**: Generic error messages for problems
//!
//! ## Related Functions
//!
//! - `LoadResPicture`: Load picture from resource file
//! - `SavePicture`: Save picture object to file
//! - `Set`: Assign object references
//! - `Nothing`: Release object references
//! - `Dir`: Check file existence before loading
//! - `App.Path`: Get application directory for relative paths

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn loadpicture_basic() {
        let source = r"
            Set Picture1.Picture = LoadPicture(filename)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_string_literal() {
        let source = r#"
            Picture1.Picture = LoadPicture("C:\Images\logo.bmp")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_if_statement() {
        let source = r"
            If fileExists Then
                Picture1.Picture = LoadPicture(imagePath)
            End If
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_concatenation() {
        let source = r#"
            Picture1.Picture = LoadPicture(App.Path & "\logo.bmp")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_empty_string() {
        let source = r#"
            Picture1.Picture = LoadPicture("")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_set_statement() {
        let source = r#"
            Set myPic = LoadPicture("image.bmp")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_form_load() {
        let source = r#"
            Private Sub Form_Load()
                Me.Picture = LoadPicture(App.Path & "\bg.bmp")
            End Sub
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_array_assignment() {
        let source = r"
            Set images(i) = LoadPicture(files(i))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_for_loop() {
        let source = r#"
            For i = 1 To 10
                Set frames(i) = LoadPicture("frame" & i & ".bmp")
            Next i
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_function_return() {
        let source = r#"
            Function GetPicture() As StdPicture
                Set GetPicture = LoadPicture("default.bmp")
            End Function
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_error_handling() {
        let source = r#"
            On Error Resume Next
            Picture1.Picture = LoadPicture(filename)
            If Err.Number <> 0 Then
                MsgBox "Error loading picture"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_with_statement() {
        let source = r#"
            With Picture1
                .Picture = LoadPicture("logo.bmp")
            End With
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_select_case() {
        let source = r#"
            Select Case imageType
                Case "logo"
                    Picture1.Picture = LoadPicture("logo.bmp")
                Case "icon"
                    Picture1.Picture = LoadPicture("icon.ico")
            End Select
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_elseif() {
        let source = r#"
            If status = "ok" Then
                imgStatus.Picture = LoadPicture("ok.ico")
            ElseIf status = "error" Then
                imgStatus.Picture = LoadPicture("error.ico")
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_parentheses() {
        let source = r"
            Set pic = (LoadPicture(filename))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_iif() {
        let source = r#"
            Picture1.Picture = IIf(enabled, LoadPicture("on.bmp"), LoadPicture("off.bmp"))
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_in_class() {
        let source = r#"
            Private Sub Class_Initialize()
                Set m_defaultPic = LoadPicture("default.bmp")
            End Sub
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_function_argument() {
        let source = r#"
            Call SetPicture(LoadPicture("image.bmp"))
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_property_assignment() {
        let source = r#"
            Set MyForm.Picture = LoadPicture("background.bmp")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_is_nothing() {
        let source = r#"
            Set pic = LoadPicture(filename)
            If pic Is Nothing Then
                MsgBox "Failed to load"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_while_wend() {
        let source = r"
            While index < maxImages
                Set images(index) = LoadPicture(files(index))
                index = index + 1
            Wend
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_do_while() {
        let source = r"
            Do While hasMore
                Set currentPic = LoadPicture(GetNextFile())
                ProcessPicture currentPic
            Loop
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_do_until() {
        let source = r"
            Do Until loaded
                On Error Resume Next
                Set pic = LoadPicture(filename)
                loaded = (Err.Number = 0)
            Loop
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_dir_check() {
        let source = r#"
            If Dir(picPath) <> "" Then
                Picture1.Picture = LoadPicture(picPath)
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_no_arguments() {
        let source = r"
            Set Picture1.Picture = LoadPicture()
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_icon_parameters() {
        let source = r#"
            Set pic = LoadPicture("icon.ico", vbLPSmall, vbLPColor)
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn loadpicture_msgbox_concatenation() {
        let source = r#"
            MsgBox "Loading: " & filename
            Picture1.Picture = LoadPicture(filename)
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/resources/loadpicture",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
