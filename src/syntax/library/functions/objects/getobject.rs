//! `GetObject` Function
//!
//! Returns a reference to an `ActiveX` object from a file or a running instance of an object.
//!
//! # Syntax
//!
//! ```vb
//! GetObject([pathname] [, class])
//! ```
//!
//! # Parameters
//!
//! - `pathname` - Optional. `String` expression that specifies the full path and name of the file containing the object to retrieve. If pathname is omitted, class is required.
//! - `class` - Optional. `String` expression that specifies the class of the object. The class argument uses the syntax `appname.objecttype` and has these parts:
//!   - `appname` - Required. The name of the application providing the object.
//!   - `objecttype` - Required. The type or class of object to create.
//!
//! # Return Value
//!
//! Returns an `Object` reference to the specified `ActiveX` object. The specific type depends on the class requested.
//!
//! # Remarks
//!
//! - Use `GetObject` to access an existing `ActiveX` object from a file or to get a reference to a running instance of an application.
//! - If pathname is omitted, `GetObject` returns a currently active object of the specified class.
//! - If no instance of the object is running, an error occurs when pathname is omitted.
//! - Some applications allow you to activate part of a file (e.g., `Excel` can activate a range in a workbook).
//! - Use the `!` character in pathname to separate the file name from the part you want to activate: `"C:\MyDoc.xls!Sheet1!R1C1:R5C5"`.
//! - `GetObject` is useful when there is a current instance of the object or if you want to create the object with a file already loaded.
//! - If there is no current instance and you don't want the object started with a file loaded, use `CreateObject`.
//! - Once an object has been activated, you reference it in code using the object variable you defined.
//! - `GetObject` always returns a single instance. If you call `GetObject` multiple times, you may get different instances.
//! - The object must support Automation for `GetObject` to work.
//!
//! # Typical Uses
//!
//! - Opening existing Office documents (`Excel`, `Word`, `PowerPoint`)
//! - Getting references to running application instances
//! - Accessing specific portions of files (`Excel` ranges, `Word` bookmarks)
//! - Working with embedded or linked `OLE` objects
//! - Automation of existing application instances
//! - Document manipulation and data extraction
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Get reference to Excel object from file
//! Dim xlApp As Object
//! Set xlApp = GetObject("C:\Reports\Sales.xls")
//!
//! ' Activate Excel
//! xlApp.Application.Visible = True
//!
//! ' Get existing Excel instance
//! Dim excelApp As Object
//! Set excelApp = GetObject(, "Excel.Application")
//!
//! ' Get specific Excel range
//! Dim xlRange As Object
//! Set xlRange = GetObject("C:\Data\Report.xls!Sheet1!R1C1:R10C5")
//!
//! ' Get Word document
//! Dim wordDoc As Object
//! Set wordDoc = GetObject("C:\Documents\Letter.doc")
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Get or Create Pattern
//!
//! ```vb
//! Function GetExcelInstance() As Object
//!     On Error Resume Next
//!     
//!     ' Try to get existing instance
//!     Set GetExcelInstance = GetObject(, "Excel.Application")
//!     
//!     If Err.Number <> 0 Then
//!         ' No instance running, create new one
//!         Set GetExcelInstance = CreateObject("Excel.Application")
//!     End If
//!     
//!     On Error GoTo 0
//! End Function
//!
//! ' Usage
//! Dim excel As Object
//! Set excel = GetExcelInstance()
//! excel.Visible = True
//! ```
//!
//! ## 2. Open Existing Excel File
//!
//! ```vb
//! Sub ProcessExcelFile(filePath As String)
//!     Dim xlApp As Object
//!     Dim xlWorkbook As Object
//!     Dim xlSheet As Object
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Open the Excel file
//!     Set xlWorkbook = GetObject(filePath)
//!     Set xlApp = xlWorkbook.Application
//!     
//!     ' Make Excel visible
//!     xlApp.Visible = True
//!     
//!     ' Access first worksheet
//!     Set xlSheet = xlWorkbook.Worksheets(1)
//!     
//!     ' Process data
//!     Debug.Print xlSheet.Range("A1").Value
//!     
//!     ' Cleanup
//!     xlWorkbook.Close SaveChanges:=False
//!     Set xlSheet = Nothing
//!     Set xlWorkbook = Nothing
//!     Set xlApp = Nothing
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error: " & Err.Description
//! End Sub
//! ```
//!
//! ## 3. Access Specific Excel Range
//!
//! ```vb
//! Sub ReadExcelRange()
//!     Dim xlRange As Object
//!     Dim cell As Variant
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Get specific range from Excel file
//!     Set xlRange = GetObject("C:\Data\Sales.xls!Sheet1!R1C1:R10C3")
//!     
//!     ' Loop through cells
//!     For Each cell In xlRange.Cells
//!         Debug.Print cell.Value
//!     Next cell
//!     
//!     ' Cleanup
//!     xlRange.Parent.Parent.Close SaveChanges:=False
//!     Set xlRange = Nothing
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error reading range: " & Err.Description
//! End Sub
//! ```
//!
//! ## 4. Connect to Running Application
//!
//! ```vb
//! Function ConnectToRunningWord() As Object
//!     On Error Resume Next
//!     
//!     Set ConnectToRunningWord = GetObject(, "Word.Application")
//!     
//!     If Err.Number <> 0 Then
//!         MsgBox "Word is not currently running"
//!         Set ConnectToRunningWord = Nothing
//!     End If
//!     
//!     On Error GoTo 0
//! End Function
//!
//! ' Usage
//! Sub UseRunningWord()
//!     Dim wordApp As Object
//!     
//!     Set wordApp = ConnectToRunningWord()
//!     
//!     If Not wordApp Is Nothing Then
//!         Debug.Print "Word has " & wordApp.Documents.Count & " documents open"
//!         Set wordApp = Nothing
//!     End If
//! End Sub
//! ```
//!
//! ## 5. Multiple File Processing
//!
//! ```vb
//! Sub ProcessMultipleExcelFiles()
//!     Dim files() As String
//!     Dim i As Long
//!     Dim xlWorkbook As Object
//!     Dim xlApp As Object
//!     
//!     files = Array("C:\Data\Jan.xls", "C:\Data\Feb.xls", "C:\Data\Mar.xls")
//!     
//!     For i = LBound(files) To UBound(files)
//!         On Error Resume Next
//!         
//!         Set xlWorkbook = GetObject(files(i))
//!         
//!         If Err.Number = 0 Then
//!             Set xlApp = xlWorkbook.Application
//!             
//!             ' Process workbook
//!             Debug.Print "Processing: " & xlWorkbook.Name
//!             Debug.Print "Sheets: " & xlWorkbook.Worksheets.Count
//!             
//!             ' Close without saving
//!             xlWorkbook.Close SaveChanges:=False
//!         Else
//!             Debug.Print "Failed to open: " & files(i)
//!         End If
//!         
//!         Set xlWorkbook = Nothing
//!         Set xlApp = Nothing
//!         
//!         On Error GoTo 0
//!     Next i
//! End Sub
//! ```
//!
//! ## 6. Word Document Automation
//!
//! ```vb
//! Sub ModifyWordDocument(filePath As String)
//!     Dim wordDoc As Object
//!     Dim wordApp As Object
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Open existing Word document
//!     Set wordDoc = GetObject(filePath)
//!     Set wordApp = wordDoc.Application
//!     
//!     ' Make Word visible
//!     wordApp.Visible = True
//!     
//!     ' Modify document
//!     wordDoc.Content.InsertAfter vbCrLf & "Added text: " & Now
//!     
//!     ' Save and close
//!     wordDoc.Save
//!     wordDoc.Close
//!     
//!     ' Cleanup
//!     Set wordDoc = Nothing
//!     Set wordApp = Nothing
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error modifying document: " & Err.Description
//! End Sub
//! ```
//!
//! ## 7. Excel Data Extraction
//!
//! ```vb
//! Function ExtractExcelData(filePath As String, _
//!                          sheetName As String, _
//!                          rangeName As String) As Variant
//!     Dim xlWorkbook As Object
//!     Dim xlSheet As Object
//!     Dim data As Variant
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     Set xlWorkbook = GetObject(filePath)
//!     Set xlSheet = xlWorkbook.Worksheets(sheetName)
//!     
//!     ' Get data from range
//!     data = xlSheet.Range(rangeName).Value
//!     
//!     ' Close workbook
//!     xlWorkbook.Close SaveChanges:=False
//!     
//!     ' Return data
//!     ExtractExcelData = data
//!     
//!     ' Cleanup
//!     Set xlSheet = Nothing
//!     Set xlWorkbook = Nothing
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     ExtractExcelData = Null
//!     If Not xlWorkbook Is Nothing Then xlWorkbook.Close SaveChanges:=False
//! End Function
//!
//! ' Usage
//! Sub DisplayData()
//!     Dim data As Variant
//!     
//!     data = ExtractExcelData("C:\Reports\Sales.xls", "Summary", "A1:C10")
//!     
//!     If Not IsNull(data) Then
//!         Debug.Print "Data extracted: " & UBound(data, 1) & " rows"
//!     End If
//! End Sub
//! ```
//!
//! ## 8. Application Instance Manager
//!
//! ```vb
//! Type AppInstance
//!     AppName As String
//!     ProgID As String
//!     IsRunning As Boolean
//!     Instance As Object
//! End Type
//!
//! Function CheckAppInstance(progID As String) As AppInstance
//!     Dim result As AppInstance
//!     
//!     result.ProgID = progID
//!     result.AppName = Split(progID, ".")(0)
//!     
//!     On Error Resume Next
//!     Set result.Instance = GetObject(, progID)
//!     On Error GoTo 0
//!     
//!     result.IsRunning = Not (result.Instance Is Nothing)
//!     
//!     CheckAppInstance = result
//! End Function
//!
//! Sub ReportRunningApplications()
//!     Dim apps() As String
//!     Dim i As Long
//!     Dim instance As AppInstance
//!     
//!     apps = Array("Excel.Application", "Word.Application", _
//!                  "PowerPoint.Application", "Outlook.Application")
//!     
//!     For i = LBound(apps) To UBound(apps)
//!         instance = CheckAppInstance(apps(i))
//!         
//!         If instance.IsRunning Then
//!             Debug.Print instance.AppName & " is running"
//!         Else
//!             Debug.Print instance.AppName & " is not running"
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ## 9. Safe Object Retrieval
//!
//! ```vb
//! Function SafeGetObject(Optional filePath As Variant, _
//!                       Optional progID As Variant) As Object
//!     Dim obj As Object
//!     
//!     On Error Resume Next
//!     
//!     If IsMissing(filePath) And IsMissing(progID) Then
//!         ' Invalid call
//!         Set SafeGetObject = Nothing
//!         Exit Function
//!     End If
//!     
//!     If IsMissing(filePath) Then
//!         ' Get running instance
//!         Set obj = GetObject(, CStr(progID))
//!     ElseIf IsMissing(progID) Then
//!         ' Get from file
//!         Set obj = GetObject(CStr(filePath))
//!     Else
//!         ' Get from file with specific class
//!         Set obj = GetObject(CStr(filePath), CStr(progID))
//!     End If
//!     
//!     If Err.Number <> 0 Then
//!         Debug.Print "GetObject failed: " & Err.Description
//!         Set obj = Nothing
//!     End If
//!     
//!     On Error GoTo 0
//!     
//!     Set SafeGetObject = obj
//! End Function
//! ```
//!
//! ## 10. Document Comparison
//!
//! ```vb
//! Function CompareExcelFiles(file1 As String, file2 As String) As Boolean
//!     Dim wb1 As Object
//!     Dim wb2 As Object
//!     Dim sheet1 As Object
//!     Dim sheet2 As Object
//!     Dim identical As Boolean
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     Set wb1 = GetObject(file1)
//!     Set wb2 = GetObject(file2)
//!     
//!     ' Compare sheet counts
//!     If wb1.Worksheets.Count <> wb2.Worksheets.Count Then
//!         CompareExcelFiles = False
//!         GoTo Cleanup
//!     End If
//!     
//!     ' Compare first sheet data
//!     Set sheet1 = wb1.Worksheets(1)
//!     Set sheet2 = wb2.Worksheets(1)
//!     
//!     If sheet1.UsedRange.Address = sheet2.UsedRange.Address Then
//!         identical = True
//!     Else
//!         identical = False
//!     End If
//!     
//!     CompareExcelFiles = identical
//!     
//! Cleanup:
//!     wb1.Close SaveChanges:=False
//!     wb2.Close SaveChanges:=False
//!     Set sheet1 = Nothing
//!     Set sheet2 = Nothing
//!     Set wb1 = Nothing
//!     Set wb2 = Nothing
//!     Exit Function
//!     
//! ErrorHandler:
//!     CompareExcelFiles = False
//!     If Not wb1 Is Nothing Then wb1.Close SaveChanges:=False
//!     If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
//! End Function
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. Document Manager Class
//!
//! ```vb
//! ' Class: DocumentManager
//! Private m_FilePath As String
//! Private m_Document As Object
//! Private m_Application As Object
//! Private m_DocumentType As String
//!
//! Public Sub OpenDocument(filePath As String)
//!     On Error GoTo ErrorHandler
//!     
//!     m_FilePath = filePath
//!     Set m_Document = GetObject(filePath)
//!     Set m_Application = m_Document.Application
//!     
//!     ' Determine document type
//!     m_DocumentType = TypeName(m_Document)
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     Err.Raise vbObjectError + 1000, , "Failed to open: " & filePath
//! End Sub
//!
//! Public Property Get IsOpen() As Boolean
//!     IsOpen = Not (m_Document Is Nothing)
//! End Property
//!
//! Public Property Get DocumentType() As String
//!     DocumentType = m_DocumentType
//! End Property
//!
//! Public Sub MakeVisible()
//!     If Not m_Application Is Nothing Then
//!         m_Application.Visible = True
//!     End If
//! End Sub
//!
//! Public Function GetProperty(propertyName As String) As Variant
//!     On Error Resume Next
//!     GetProperty = CallByName(m_Document, propertyName, VbGet)
//! End Function
//!
//! Public Sub CloseDocument(Optional saveChanges As Boolean = False)
//!     If Not m_Document Is Nothing Then
//!         m_Document.Close saveChanges
//!         Set m_Document = Nothing
//!         Set m_Application = Nothing
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseDocument False
//! End Sub
//! ```
//!
//! ## 2. Batch File Processor
//!
//! ```vb
//! Type ProcessingResult
//!     FileName As String
//!     Success As Boolean
//!     ErrorMessage As String
//!     ProcessedDate As Date
//! End Type
//!
//! Function BatchProcessFiles(files As Collection, _
//!                           processingFunc As String) As Collection
//!     Dim results As New Collection
//!     Dim file As Variant
//!     Dim doc As Object
//!     Dim result As ProcessingResult
//!     
//!     For Each file In files
//!         result.FileName = CStr(file)
//!         result.ProcessedDate = Now
//!         
//!         On Error Resume Next
//!         Set doc = GetObject(CStr(file))
//!         
//!         If Err.Number = 0 Then
//!             ' Call custom processing function
//!             Application.Run processingFunc, doc
//!             
//!             doc.Save
//!             doc.Close
//!             
//!             result.Success = True
//!             result.ErrorMessage = ""
//!         Else
//!             result.Success = False
//!             result.ErrorMessage = Err.Description
//!         End If
//!         
//!         On Error GoTo 0
//!         
//!         Set doc = Nothing
//!         results.Add result
//!     Next file
//!     
//!     Set BatchProcessFiles = results
//! End Function
//! ```
//!
//! ## 3. Smart Office Connector
//!
//! ```vb
//! Class OfficeConnector
//!     Private m_Instances As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_Instances = New Collection
//!     End Sub
//!     
//!     Public Function GetOrCreateExcel() As Object
//!         Dim excel As Object
//!         Dim key As String
//!         
//!         key = "Excel.Application"
//!         
//!         ' Check cache
//!         On Error Resume Next
//!         Set excel = m_Instances(key)
//!         On Error GoTo 0
//!         
//!         If excel Is Nothing Then
//!             ' Try to get existing instance
//!             On Error Resume Next
//!             Set excel = GetObject(, key)
//!             On Error GoTo 0
//!             
//!             If excel Is Nothing Then
//!                 ' Create new instance
//!                 Set excel = CreateObject(key)
//!             End If
//!             
//!             ' Cache instance
//!             m_Instances.Add excel, key
//!         End If
//!         
//!         Set GetOrCreateExcel = excel
//!     End Function
//!     
//!     Public Function OpenFile(filePath As String) As Object
//!         Dim doc As Object
//!         
//!         On Error GoTo ErrorHandler
//!         
//!         Set doc = GetObject(filePath)
//!         Set OpenFile = doc
//!         
//!         Exit Function
//!         
//!     ErrorHandler:
//!         Set OpenFile = Nothing
//!     End Function
//!     
//!     Public Sub CloseAll()
//!         Dim item As Variant
//!         
//!         For Each item In m_Instances
//!             On Error Resume Next
//!             item.Quit
//!             On Error GoTo 0
//!         Next item
//!         
//!         Set m_Instances = New Collection
//!     End Sub
//!     
//!     Private Sub Class_Terminate()
//!         CloseAll
//!     End Sub
//! End Class
//! ```
//!
//! ## 4. Document Cache System
//!
//! ```vb
//! Type CachedDocument
//!     FilePath As String
//!     Document As Object
//!     LastAccessed As Date
//!     AccessCount As Long
//! End Type
//!
//! Private m_DocumentCache As Collection
//! Private Const CACHE_TIMEOUT = 300 ' 5 minutes in seconds
//!
//! Sub InitializeDocumentCache()
//!     Set m_DocumentCache = New Collection
//! End Sub
//!
//! Function GetCachedDocument(filePath As String) As Object
//!     Dim cached As CachedDocument
//!     Dim i As Long
//!     Dim found As Boolean
//!     
//!     ' Search cache
//!     For i = 1 To m_DocumentCache.Count
//!         cached = m_DocumentCache(i)
//!         
//!         If cached.FilePath = filePath Then
//!             ' Check if cache is still valid
//!             If DateDiff("s", cached.LastAccessed, Now) < CACHE_TIMEOUT Then
//!                 cached.LastAccessed = Now
//!                 cached.AccessCount = cached.AccessCount + 1
//!                 m_DocumentCache.Remove i
//!                 m_DocumentCache.Add cached, filePath
//!                 
//!                 Set GetCachedDocument = cached.Document
//!                 Exit Function
//!             Else
//!                 ' Cache expired, remove it
//!                 cached.Document.Close SaveChanges:=False
//!                 m_DocumentCache.Remove i
//!                 Exit For
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Not in cache, open new
//!     On Error Resume Next
//!     Set cached.Document = GetObject(filePath)
//!     On Error GoTo 0
//!     
//!     If Not cached.Document Is Nothing Then
//!         cached.FilePath = filePath
//!         cached.LastAccessed = Now
//!         cached.AccessCount = 1
//!         
//!         m_DocumentCache.Add cached, filePath
//!         Set GetCachedDocument = cached.Document
//!     Else
//!         Set GetCachedDocument = Nothing
//!     End If
//! End Function
//!
//! Sub ClearDocumentCache()
//!     Dim cached As CachedDocument
//!     Dim i As Long
//!     
//!     For i = m_DocumentCache.Count To 1 Step -1
//!         cached = m_DocumentCache(i)
//!         cached.Document.Close SaveChanges:=False
//!         m_DocumentCache.Remove i
//!     Next i
//! End Sub
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafelyGetObject(Optional pathName As String, _
//!                         Optional className As String) As Object
//!     On Error GoTo ErrorHandler
//!     
//!     If pathName <> "" And className <> "" Then
//!         Set SafelyGetObject = GetObject(pathName, className)
//!     ElseIf pathName <> "" Then
//!         Set SafelyGetObject = GetObject(pathName)
//!     ElseIf className <> "" Then
//!         Set SafelyGetObject = GetObject(, className)
//!     Else
//!         Set SafelyGetObject = Nothing
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 429  ' ActiveX component can't create object
//!             Debug.Print "Object unavailable or not registered"
//!         Case 432  ' File name or class name not found
//!             Debug.Print "File or class not found: " & pathName & " " & className
//!         Case 462  ' Remote server machine does not exist
//!             Debug.Print "Server not available"
//!         Case 70   ' Permission denied
//!             Debug.Print "Access denied"
//!         Case Else
//!             Debug.Print "Error " & Err.Number & ": " & Err.Description
//!     End Select
//!     
//!     Set SafelyGetObject = Nothing
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 429**: `ActiveX` component can't create object - object not available or not registered.
//! - **Error 432**: File name or class name not found during Automation operation.
//! - **Error 462**: Remote server machine does not exist or is unavailable.
//! - **Error 70**: Permission denied - file is locked or insufficient permissions.
//! - **Error 5**: Invalid procedure call - incorrect parameters.
//!
//! # Performance Considerations
//!
//! - `GetObject` can be slower than `CreateObject` for new instances
//! - Opening files with `GetObject` loads the entire file into memory
//! - Use specific ranges when possible to minimize memory usage
//! - Consider caching object references for frequently accessed files
//! - Close objects when done to free resources
//! - For batch operations, reuse application instances
//!
//! # Best Practices
//!
//! 1. **Always use error handling** - files may not exist or be accessible
//! 2. **Close objects explicitly** - release resources promptly
//! 3. **Set object variables to Nothing** - ensure cleanup
//! 4. **Check if object is Nothing** before using
//! 5. **Use specific object types** when possible (late vs early binding)
//! 6. **Handle both file and instance retrieval** scenarios
//! 7. **Cache references** for frequently accessed objects
//! 8. **Test file existence** before calling `GetObject`
//!
//! # Comparison with Other Functions
//!
//! ## `GetObject` vs `CreateObject`
//!
//! ```vb
//! ' GetObject - Get existing instance or open file
//! Set excel = GetObject(, "Excel.Application")  ' Gets running instance
//! Set workbook = GetObject("C:\Data.xls")       ' Opens existing file
//!
//! ' CreateObject - Always creates new instance
//! Set excel = CreateObject("Excel.Application") ' Creates new instance
//! Set workbook = excel.Workbooks.Open("C:\Data.xls") ' Opens file
//! ```
//!
//! ## `GetObject` with File vs Without
//!
//! ```vb
//! ' With file - Opens the file
//! Set doc = GetObject("C:\Report.xls")
//!
//! ' Without file - Gets running instance
//! Set app = GetObject(, "Excel.Application")
//!
//! ' With both - Opens file with specific application
//! Set doc = GetObject("C:\Data.txt", "Excel.Application")
//! ```
//!
//! # Limitations
//!
//! - Requires the object to support Automation
//! - File must exist for pathname-based calls
//! - Application must be registered on the system
//! - May fail if file is already open exclusively
//! - Limited control over how the file is opened
//! - Cannot specify detailed options (read-only, etc.)
//! - Not all file types support the `!` notation for partial activation
//! - May behave differently across Office versions
//!
//! # File Activation Syntax
//!
//! For files that support partial activation:
//!
//! ```vb
//! ' Excel - specific range
//! Set range = GetObject("C:\Data.xls!Sheet1!R1C1:R10C10")
//!
//! ' Excel - named range
//! Set range = GetObject("C:\Data.xls!MyRange")
//!
//! ' Word - bookmark (if supported)
//! Set bookmark = GetObject("C:\Doc.doc!MyBookmark")
//! ```
//!
//! # Related Functions
//!
//! - `CreateObject` - Creates a new instance of an `ActiveX` object
//! - `GetAutoServerSettings` - Returns `DCOM` server security settings
//! - `CallByName` - Calls a method or accesses a property dynamically
//! - `TypeName` - Returns type information about an object
//! - `IsObject` - Checks if a variable contains an object reference
//! - `Set` - Assigns an object reference to a variable

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn getobject_basic() {
        let source = r#"Set xlApp = GetObject("C:\Reports\Sales.xls")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_with_class() {
        let source = r#"Set app = GetObject(, "Excel.Application")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_both_params() {
        let source = r"Set doc = GetObject(filePath, className)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_in_function() {
        let source = r#"Function GetExcelInstance() As Object
    Set GetExcelInstance = GetObject(, "Excel.Application")
End Function"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_error_handling() {
        let source = r#"On Error Resume Next
Set obj = GetObject("C:\data.xls")
On Error GoTo 0"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_if_statement() {
        let source = r#"If Not GetObject(, "Excel.Application") Is Nothing Then
    MsgBox "Excel is running"
End If"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_range_notation() {
        let source = r#"Set xlRange = GetObject("C:\Data\Report.xls!Sheet1!R1C1:R10C5")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_word_document() {
        let source = r#"Set wordDoc = GetObject("C:\Documents\Letter.doc")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_for_loop() {
        let source = r"For i = 1 To fileCount
    Set doc = GetObject(files(i))
Next i";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_is_nothing_check() {
        let source = r"If GetObject(, progID) Is Nothing Then Exit Sub";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_with_application() {
        let source = r"Set xlApp = GetObject(filePath).Application";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_select_case() {
        let source = r#"Select Case TypeName(GetObject(filePath))
    Case "Workbook"
        Debug.Print "Excel file"
End Select"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_do_loop() {
        let source = r#"Do Until Not GetObject(, "Excel.Application") Is Nothing
    DoEvents
Loop"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_class_member() {
        let source = r"Set m_Document = GetObject(m_FilePath)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_type_field() {
        let source = r"Set cached.Document = GetObject(filePath)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_collection_add() {
        let source = r#"m_Instances.Add GetObject(, "Excel.Application"), "Excel""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_with_statement() {
        let source = r#"With GetObject("C:\data.xls")
    .Application.Visible = True
End With"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_debug_print() {
        let source = r#"Debug.Print "Count: " & GetObject(, "Excel.Application").Workbooks.Count"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_msgbox() {
        let source = r"MsgBox TypeName(GetObject(filePath))";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_property() {
        let source = r"Property Get CurrentDocument() As Object
    Set CurrentDocument = GetObject(m_FilePath)
End Property";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_concatenation() {
        let source = r#"filePath = "C:\Data\" & fileName & ".xls"
Set doc = GetObject(filePath)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_for_each() {
        let source = r"For Each file In files
    Set doc = GetObject(CStr(file))
Next file";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_call_by_name() {
        let source = r#"result = CallByName(GetObject(filePath), "Save", VbMethod)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_method_call() {
        let source = r"GetObject(filePath).Save";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_nested_property() {
        let source = r#"value = GetObject(filePath).Worksheets(1).Range("A1").Value"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_parentheses() {
        let source = r#"Set app = (GetObject(, "Excel.Application"))"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn getobject_iif() {
        let source = r"Set obj = IIf(condition, GetObject(file1), GetObject(file2))";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/getobject",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
