//! # `CreateObject` Function
//!
//! Creates and returns a reference to an `ActiveX` object (`COM` object).
//!
//! ## Syntax
//!
//! ```vb
//! CreateObject(class, [servername])
//! ```
//!
//! ## Parameters
//!
//! - **`class`**: Required. `String` expression representing the programmatic identifier (`ProgID`) of
//!   the object to create. The format is typically "Application.ObjectType" or
//!   "Library.Class".
//!
//! - **`servername`**: Optional. `String` expression representing the name of the network server where
//!   the object will be created. If omitted or an empty string (""), the object is created on the
//!   local machine. This parameter is only used for `DCOM` (`Distributed COM`).
//!
//! ## Return Value
//!
//! Returns an `Object` reference to the created `COM` object. The actual type depends on the class
//! specified. Returns `Nothing` if the object cannot be created.
//!
//! ## Remarks
//!
//! `CreateObject` is used to instantiate `COM` objects at runtime. This is known as late binding,
//! as opposed to early binding where you reference the object library and declare objects with
//! specific types at design time.
//!
//! **Important Characteristics:**
//!
//! - Creates objects using late binding (runtime resolution)
//! - Requires the `COM` object to be registered on the system
//! - Returns generic `Object` type (requires type casting for `IntelliSense`)
//! - Slower than early binding but more flexible
//! - No compile-time type checking
//! - Enables automation of external applications
//! - Can create objects on remote servers (`DCOM`)
//!
//! ## Common `ProgID`s
//!
//! | `ProgID` | Description |
//! |--------|-------------|
//! | "Excel.Application" | Microsoft Excel application |
//! | "Word.Application" | Microsoft Word application |
//! | "Scripting.FileSystemObject" | File system object for file operations |
//! | "Scripting.Dictionary" | Dictionary object for key-value pairs |
//! | "ADODB.Connection" | ADO database connection |
//! | "ADODB.Recordset" | ADO recordset for database queries |
//! | "Shell.Application" | Windows Shell automation |
//! | "WScript.Shell" | Windows Script Host Shell object |
//! | "MSXML2.DOMDocument" | XML DOM parser |
//! | "CDO.Message" | Collaboration Data Objects for email |
//! | "InternetExplorer.Application" | Internet Explorer automation |
//! | "Outlook.Application" | Microsoft Outlook application |
//! | "Access.Application" | Microsoft Access application |
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Create an Excel application object
//! Dim xlApp As Object
//! Set xlApp = CreateObject("Excel.Application")
//! xlApp.Visible = True
//! xlApp.Workbooks.Add
//! Set xlApp = Nothing
//!
//! ' Create a FileSystemObject
//! Dim fso As Object
//! Set fso = CreateObject("Scripting.FileSystemObject")
//! ```
//!
//! ### Microsoft Excel Automation
//!
//! ```vb
//! Sub CreateExcelSpreadsheet()
//!     Dim xlApp As Object
//!     Dim xlBook As Object
//!     Dim xlSheet As Object
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Create Excel application
//!     Set xlApp = CreateObject("Excel.Application")
//!     xlApp.Visible = True
//!     
//!     ' Add a workbook
//!     Set xlBook = xlApp.Workbooks.Add
//!     Set xlSheet = xlBook.Worksheets(1)
//!     
//!     ' Add data
//!     xlSheet.Cells(1, 1).Value = "Name"
//!     xlSheet.Cells(1, 2).Value = "Value"
//!     xlSheet.Cells(2, 1).Value = "Item 1"
//!     xlSheet.Cells(2, 2).Value = 100
//!     
//!     ' Clean up
//!     Set xlSheet = Nothing
//!     Set xlBook = Nothing
//!     xlApp.Quit
//!     Set xlApp = Nothing
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error: " & Err.Description
//!     If Not xlApp Is Nothing Then
//!         xlApp.Quit
//!         Set xlApp = Nothing
//!     End If
//! End Sub
//! ```
//!
//! ### File System Operations
//!
//! ```vb
//! Function FileExists(filePath As String) As Boolean
//!     Dim fso As Object
//!     Set fso = CreateObject("Scripting.FileSystemObject")
//!     FileExists = fso.FileExists(filePath)
//!     Set fso = Nothing
//! End Function
//!
//! Function GetFileSize(filePath As String) As Long
//!     Dim fso As Object
//!     Dim file As Object
//!     
//!     Set fso = CreateObject("Scripting.FileSystemObject")
//!     If fso.FileExists(filePath) Then
//!         Set file = fso.GetFile(filePath)
//!         GetFileSize = file.Size
//!         Set file = Nothing
//!     End If
//!     Set fso = Nothing
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Dictionary for Key-Value Storage
//!
//! ```vb
//! Function CreateDictionary() As Object
//!     Dim dict As Object
//!     Set dict = CreateObject("Scripting.Dictionary")
//!     
//!     ' Add items
//!     dict.Add "Name", "John Doe"
//!     dict.Add "Age", 30
//!     dict.Add "City", "New York"
//!     
//!     Set CreateDictionary = dict
//! End Function
//! ```
//!
//! ### Database Connection
//!
//! ```vb
//! Function OpenDatabase(connectionString As String) As Object
//!     Dim conn As Object
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     Set conn = CreateObject("ADODB.Connection")
//!     conn.Open connectionString
//!     
//!     Set OpenDatabase = conn
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Database error: " & Err.Description
//!     Set OpenDatabase = Nothing
//! End Function
//! ```
//!
//! ### XML Document Processing
//!
//! ```vb
//! Function LoadXMLFile(filePath As String) As Object
//!     Dim xmlDoc As Object
//!     
//!     Set xmlDoc = CreateObject("MSXML2.DOMDocument")
//!     xmlDoc.async = False
//!     
//!     If xmlDoc.Load(filePath) Then
//!         Set LoadXMLFile = xmlDoc
//!     Else
//!         MsgBox "Error loading XML: " & xmlDoc.parseError.reason
//!         Set LoadXMLFile = Nothing
//!     End If
//! End Function
//! ```
//!
//! ### Shell Commands
//!
//! ```vb
//! Sub RunCommand(command As String)
//!     Dim shell As Object
//!     Set shell = CreateObject("WScript.Shell")
//!     shell.Run command, 1, True  ' Wait for completion
//!     Set shell = Nothing
//! End Sub
//!
//! Function GetEnvironmentVariable(varName As String) As String
//!     Dim shell As Object
//!     Set shell = CreateObject("WScript.Shell")
//!     GetEnvironmentVariable = shell.ExpandEnvironmentStrings("%" & varName & "%")
//!     Set shell = Nothing
//! End Function
//! ```
//!
//! ### Email Sending (CDO)
//!
//! ```vb
//! Sub SendEmail(toAddr As String, subject As String, body As String)
//!     Dim msg As Object
//!     Dim config As Object
//!     
//!     Set msg = CreateObject("CDO.Message")
//!     Set config = CreateObject("CDO.Configuration")
//!     
//!     ' Configure SMTP settings
//!     With config.Fields
//!         .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
//!         .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.example.com"
//!         .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
//!         .Update
//!     End With
//!     
//!     ' Send message
//!     With msg
//!         Set .Configuration = config
//!         .To = toAddr
//!         .From = "sender@example.com"
//!         .Subject = subject
//!         .TextBody = body
//!         .Send
//!     End With
//!     
//!     Set msg = Nothing
//!     Set config = Nothing
//! End Sub
//! ```
//!
//! ### Word Document Creation
//!
//! ```vb
//! Sub CreateWordDocument()
//!     Dim wordApp As Object
//!     Dim wordDoc As Object
//!     
//!     Set wordApp = CreateObject("Word.Application")
//!     wordApp.Visible = True
//!     
//!     Set wordDoc = wordApp.Documents.Add
//!     wordDoc.Content.Text = "This is a test document."
//!     
//!     Set wordDoc = Nothing
//!     Set wordApp = Nothing
//! End Sub
//! ```
//!
//! ### Registry Access
//!
//! ```vb
//! Function ReadRegistry(keyPath As String) As String
//!     Dim shell As Object
//!     Set shell = CreateObject("WScript.Shell")
//!     
//!     On Error Resume Next
//!     ReadRegistry = shell.RegRead(keyPath)
//!     
//!     Set shell = Nothing
//! End Function
//!
//! Sub WriteRegistry(keyPath As String, value As String)
//!     Dim shell As Object
//!     Set shell = CreateObject("WScript.Shell")
//!     shell.RegWrite keyPath, value
//!     Set shell = Nothing
//! End Sub
//! ```
//!
//! ### HTTP Request
//!
//! ```vb
//! Function GetWebPage(url As String) As String
//!     Dim http As Object
//!     
//!     Set http = CreateObject("MSXML2.XMLHTTP")
//!     http.Open "GET", url, False
//!     http.Send
//!     
//!     If http.Status = 200 Then
//!         GetWebPage = http.responseText
//!     End If
//!     
//!     Set http = Nothing
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Remote Object Creation (DCOM)
//!
//! ```vb
//! Sub CreateRemoteObject()
//!     Dim obj As Object
//!     
//!     ' Create object on remote server
//!     Set obj = CreateObject("MyApp.MyClass", "\\ServerName")
//!     
//!     ' Use the remote object
//!     obj.DoSomething
//!     
//!     Set obj = Nothing
//! End Sub
//! ```
//!
//! ### Object Factory Pattern
//!
//! ```vb
//! Function CreateObjectSafe(progID As String) As Object
//!     On Error GoTo ErrorHandler
//!     
//!     Set CreateObjectSafe = CreateObject(progID)
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Failed to create object: " & progID & vbCrLf & _
//!            "Error: " & Err.Description, vbCritical
//!     Set CreateObjectSafe = Nothing
//! End Function
//! ```
//!
//! ### Version-Specific Object Creation
//!
//! ```vb
//! Function CreateExcelObject() As Object
//!     On Error Resume Next
//!     
//!     ' Try different versions in order of preference
//!     Set CreateExcelObject = CreateObject("Excel.Application.16")  ' Office 2016
//!     If CreateExcelObject Is Nothing Then
//!         Set CreateExcelObject = CreateObject("Excel.Application.15")  ' Office 2013
//!     End If
//!     If CreateExcelObject Is Nothing Then
//!         Set CreateExcelObject = CreateObject("Excel.Application")  ' Any version
//!     End If
//!     
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function CreateObjectWithErrorHandling(progID As String) As Object
//!     On Error GoTo ErrorHandler
//!     
//!     Set CreateObjectWithErrorHandling = CreateObject(progID)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 429  ' ActiveX component can't create object
//!             MsgBox "The COM object '" & progID & "' is not registered on this system.", _
//!                    vbCritical, "Object Not Found"
//!         Case 70   ' Permission denied
//!             MsgBox "Permission denied creating object: " & progID, vbCritical
//!         Case Else
//!             MsgBox "Error creating object: " & progID & vbCrLf & _
//!                    "Error " & Err.Number & ": " & Err.Description, vbCritical
//!     End Select
//!     
//!     Set CreateObjectWithErrorHandling = Nothing
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 429** (`ActiveX` component can't create object): `Object` not registered or not installed
//! - **Error 70** (Permission denied): Insufficient permissions to create the object
//! - **Error 462** (The remote server machine does not exist or is unavailable): `DCOM` server issue
//! - **Error 13** (Type mismatch): Invalid `ProgID` format
//!
//! ## Performance Considerations
//!
//! - Late binding (`CreateObject`) is slower than early binding
//! - No `IntelliSense` or compile-time checking with `CreateObject`
//! - Reuse objects when possible instead of creating multiple instances
//! - Always set objects to `Nothing` when done to release resources
//! - Creating objects on remote servers has network overhead
//!
//! ## Early Binding vs Late Binding
//!
//! ### Late Binding (`CreateObject`)
//! ```vb
//! Dim xlApp As Object  ' Generic Object type
//! Set xlApp = CreateObject("Excel.Application")
//! xlApp.Visible = True  ' No IntelliSense
//! ```
//!
//! **Advantages:**
//! - No reference needed at design time
//! - Works with any version of the COM object
//! - More flexible for distribution
//!
//! **Disadvantages:**
//! - Slower performance
//! - No `IntelliSense`
//! - No compile-time checking
//! - Errors only at runtime
//!
//! ### Early Binding (Object Library Reference)
//! ```vb
//! ' Add reference to "Microsoft Excel XX.0 Object Library"
//! Dim xlApp As Excel.Application
//! Set xlApp = New Excel.Application
//! xlApp.Visible = True  ' IntelliSense available
//! ```
//!
//! **Advantages:**
//! - Faster performance
//! - `IntelliSense` support
//! - Compile-time checking
//! - Better debugging
//!
//! **Disadvantages:**
//! - Requires reference at design time
//! - Version-specific
//! - Larger deployment package
//!
//! ## Best Practices
//!
//! ### Always Clean Up Objects
//!
//! ```vb
//! Sub ProperCleanup()
//!     Dim xlApp As Object
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     Set xlApp = CreateObject("Excel.Application")
//!     ' Use the object...
//!     
//!     ' Clean up
//!     If Not xlApp Is Nothing Then
//!         xlApp.Quit
//!         Set xlApp = Nothing
//!     End If
//!     Exit Sub
//!     
//! ErrorHandler:
//!     If Not xlApp Is Nothing Then
//!         xlApp.Quit
//!         Set xlApp = Nothing
//!     End If
//! End Sub
//! ```
//!
//! ### Check Object Creation Success
//!
//! ```vb
//! Dim obj As Object
//! Set obj = CreateObject("Some.Object")
//! If obj Is Nothing Then
//!     MsgBox "Failed to create object"
//!     Exit Sub
//! End If
//! ```
//!
//! ### Use Specific Error Handling
//!
//! ```vb
//! On Error Resume Next
//! Set obj = CreateObject("Excel.Application")
//! If Err.Number <> 0 Then
//!     MsgBox "Excel not available: " & Err.Description
//!     Exit Sub
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Limitations
//!
//! - Requires COM object to be registered on the system
//! - No compile-time type checking
//! - Slower than early binding
//! - No `IntelliSense` support in IDE
//! - `DCOM` requires proper network and security configuration
//! - Cannot create objects with parameterized constructors
//! - Limited to COM/ActiveX objects only
//!
//! ## Related Functions
//!
//! - `GetObject`: Gets reference to existing object or creates from file
//! - `New`: Creates early-bound object (requires reference)
//! - `Set`: Assigns object reference
//! - `Nothing`: Releases object reference
//! - `Is`: Compares object references
//! - `TypeName`: Returns type name of object

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn createobject_basic() {
        let source = r#"
Set obj = CreateObject("Excel.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_fso() {
        let source = r#"
Set fso = CreateObject("Scripting.FileSystemObject")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_dictionary() {
        let source = r#"
Set dict = CreateObject("Scripting.Dictionary")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_server() {
        let source = r#"
Set obj = CreateObject("MyApp.MyClass", "\\ServerName")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_adodb_connection() {
        let source = r#"
Set conn = CreateObject("ADODB.Connection")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_xml() {
        let source = r#"
Set xmlDoc = CreateObject("MSXML2.DOMDocument")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_shell() {
        let source = r#"
Set shell = CreateObject("WScript.Shell")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_in_function() {
        let source = r#"
Function GetFileSystem() As Object
    Set GetFileSystem = CreateObject("Scripting.FileSystemObject")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_error_handling() {
        let source = r#"
On Error Resume Next
Set obj = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Error"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_word() {
        let source = r#"
Set wordApp = CreateObject("Word.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_assignment() {
        let source = r#"
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_cdo_message() {
        let source = r#"
Set msg = CreateObject("CDO.Message")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_http() {
        let source = r#"
Set http = CreateObject("MSXML2.XMLHTTP")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_in_if() {
        let source = r#"
If CreateObject("Excel.Application") Is Nothing Then
    MsgBox "Failed"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_access() {
        let source = r#"
Set accApp = CreateObject("Access.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_outlook() {
        let source = r#"
Set outlookApp = CreateObject("Outlook.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_recordset() {
        let source = r#"
Set rs = CreateObject("ADODB.Recordset")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_version() {
        let source = r#"
Set xlApp = CreateObject("Excel.Application.16")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_internet_explorer() {
        let source = r#"
Set ie = CreateObject("InternetExplorer.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_in_sub() {
        let source = r#"
Sub Initialize()
    Set obj = CreateObject("Scripting.FileSystemObject")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_shell_application() {
        let source = r#"
Set shell = CreateObject("Shell.Application")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_multiple_in_function() {
        let source = r#"
Function SendEmail()
    Set msg = CreateObject("CDO.Message")
    Set config = CreateObject("CDO.Configuration")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_in_select_case() {
        let source = r#"
Select Case appType
    Case "Excel"
        Set app = CreateObject("Excel.Application")
    Case "Word"
        Set app = CreateObject("Word.Application")
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_immediate_use() {
        let source = r#"
result = CreateObject("Scripting.FileSystemObject").FileExists(path)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn createobject_with_whitespace() {
        let source = r#"
Set obj = CreateObject( "Excel.Application" )
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CreateObject"));
        assert!(debug.contains("Identifier"));
    }
}
