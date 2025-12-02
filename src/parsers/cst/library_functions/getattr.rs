//! `GetAttr` Function
//!
//! Returns an Integer representing the attributes of a file, directory, or folder.
//!
//! # Syntax
//!
//! ```vb
//! GetAttr(pathname)
//! ```
//!
//! # Parameters
//!
//! - `pathname` - Required. String expression that specifies a file name. May include directory or folder, and drive.
//!
//! # Return Value
//!
//! Returns an Integer representing the sum of the following attribute values:
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbNormal | 0 | Normal (no attributes set) |
//! | vbReadOnly | 1 | Read-only |
//! | vbHidden | 2 | Hidden |
//! | vbSystem | 4 | System file |
//! | vbVolume | 8 | Volume label |
//! | vbDirectory | 16 | Directory or folder |
//! | vbArchive | 32 | File has changed since last backup |
//! | vbAlias | 64 | Specified file name is an alias (Macintosh only) |
//!
//! # Remarks
//!
//! - Use the And operator to test for a specific attribute.
//! - The return value can be a combination of multiple attributes.
//! - To check if a file has a specific attribute, use bitwise AND comparison.
//! - `GetAttr` generates an error if the file doesn't exist.
//! - Use `Dir` to check if a file exists before calling `GetAttr`.
//! - On systems other than Macintosh, `vbAlias` is never set.
//! - The `vbVolume` constant is used for volume labels only.
//!
//! # Typical Uses
//!
//! - Checking if a file is read-only before attempting to modify it
//! - Detecting hidden or system files
//! - Verifying if a path points to a directory
//! - Determining file attributes for backup operations
//! - Filtering files based on attributes
//! - Security and permission checks
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Check if file is read-only
//! Dim attr As Integer
//! attr = GetAttr("C:\data.txt")
//!
//! If attr And vbReadOnly Then
//!     MsgBox "File is read-only"
//! End If
//!
//! ' Check if path is a directory
//! Dim pathAttr As Integer
//! pathAttr = GetAttr("C:\MyFolder")
//!
//! If pathAttr And vbDirectory Then
//!     MsgBox "This is a directory"
//! End If
//!
//! ' Check for hidden file
//! Dim fileAttr As Integer
//! fileAttr = GetAttr("C:\hidden.sys")
//!
//! If fileAttr And vbHidden Then
//!     MsgBox "File is hidden"
//! End If
//!
//! ' Check for system file
//! If GetAttr("C:\Windows\System32\ntoskrnl.exe") And vbSystem Then
//!     MsgBox "This is a system file"
//! End If
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Check Read-Only Before Writing
//!
//! ```vb
//! Sub WriteToFile(filename As String, content As String)
//!     Dim attr As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr = GetAttr(filename)
//!     
//!     If attr And vbReadOnly Then
//!         MsgBox "Cannot write to read-only file: " & filename
//!         Exit Sub
//!     End If
//!     
//!     ' Proceed with writing
//!     Dim fileNum As Integer
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     Print #fileNum, content
//!     Close #fileNum
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error accessing file: " & Err.Description
//! End Sub
//! ```
//!
//! ## 2. Detect Directory vs File
//!
//! ```vb
//! Function IsDirectory(path As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim attr As Integer
//!     attr = GetAttr(path)
//!     
//!     IsDirectory = (attr And vbDirectory) <> 0
//!     Exit Function
//!     
//! ErrorHandler:
//!     IsDirectory = False
//! End Function
//!
//! ' Usage
//! If IsDirectory("C:\Windows") Then
//!     Debug.Print "Path is a directory"
//! Else
//!     Debug.Print "Path is a file"
//! End If
//! ```
//!
//! ## 3. Find Hidden Files
//!
//! ```vb
//! Sub ListHiddenFiles(folderPath As String)
//!     Dim filename As String
//!     Dim fullPath As String
//!     Dim attr As Integer
//!     
//!     filename = Dir(folderPath & "\*.*", vbHidden)
//!     
//!     Do While filename <> ""
//!         fullPath = folderPath & "\" & filename
//!         
//!         On Error Resume Next
//!         attr = GetAttr(fullPath)
//!         On Error GoTo 0
//!         
//!         If (attr And vbHidden) And Not (attr And vbDirectory) Then
//!             Debug.Print "Hidden file: " & filename
//!         End If
//!         
//!         filename = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ## 4. Check Multiple Attributes
//!
//! ```vb
//! Function GetFileAttributeDescription(filename As String) As String
//!     Dim attr As Integer
//!     Dim description As String
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr = GetAttr(filename)
//!     description = ""
//!     
//!     If attr = vbNormal Then
//!         description = "Normal"
//!     Else
//!         If attr And vbReadOnly Then description = description & "Read-Only "
//!         If attr And vbHidden Then description = description & "Hidden "
//!         If attr And vbSystem Then description = description & "System "
//!         If attr And vbDirectory Then description = description & "Directory "
//!         If attr And vbArchive Then description = description & "Archive "
//!     End If
//!     
//!     GetFileAttributeDescription = Trim(description)
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileAttributeDescription = "Error: " & Err.Description
//! End Function
//! ```
//!
//! ## 5. Safe File Delete
//!
//! ```vb
//! Function SafeDeleteFile(filename As String) As Boolean
//!     Dim attr As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Check if file exists and get attributes
//!     attr = GetAttr(filename)
//!     
//!     ' Don't delete system or read-only files
//!     If attr And vbSystem Then
//!         MsgBox "Cannot delete system file: " & filename
//!         SafeDeleteFile = False
//!         Exit Function
//!     End If
//!     
//!     ' Remove read-only attribute if set
//!     If attr And vbReadOnly Then
//!         SetAttr filename, attr And Not vbReadOnly
//!     End If
//!     
//!     ' Delete the file
//!     Kill filename
//!     SafeDeleteFile = True
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Error deleting file: " & Err.Description
//!     SafeDeleteFile = False
//! End Function
//! ```
//!
//! ## 6. List Files by Attribute
//!
//! ```vb
//! Sub ListFilesByAttribute(folderPath As String, attributeFlag As Integer)
//!     Dim filename As String
//!     Dim fullPath As String
//!     Dim attr As Integer
//!     
//!     filename = Dir(folderPath & "\*.*", vbNormal + vbHidden + vbSystem)
//!     
//!     Do While filename <> ""
//!         fullPath = folderPath & "\" & filename
//!         
//!         On Error Resume Next
//!         attr = GetAttr(fullPath)
//!         On Error GoTo 0
//!         
//!         If (attr And attributeFlag) And Not (attr And vbDirectory) Then
//!             Debug.Print filename & " (Attribute: " & attr & ")"
//!         End If
//!         
//!         filename = Dir
//!     Loop
//! End Sub
//!
//! ' Usage
//! ListFilesByAttribute "C:\Temp", vbArchive  ' List files needing backup
//! ```
//!
//! ## 7. Backup File Scanner
//!
//! ```vb
//! Function NeedsBackup(filename As String) As Boolean
//!     Dim attr As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr = GetAttr(filename)
//!     
//!     ' Check if archive bit is set
//!     NeedsBackup = (attr And vbArchive) <> 0
//!     Exit Function
//!     
//! ErrorHandler:
//!     NeedsBackup = False
//! End Function
//!
//! Sub FindFilesNeedingBackup(folderPath As String, lst As ListBox)
//!     Dim filename As String
//!     Dim fullPath As String
//!     
//!     lst.Clear
//!     filename = Dir(folderPath & "\*.*")
//!     
//!     Do While filename <> ""
//!         fullPath = folderPath & "\" & filename
//!         
//!         If NeedsBackup(fullPath) Then
//!             lst.AddItem filename
//!         End If
//!         
//!         filename = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ## 8. File Permission Check
//!
//! ```vb
//! Function CanModifyFile(filename As String) As Boolean
//!     Dim attr As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr = GetAttr(filename)
//!     
//!     ' File can be modified if it's not read-only and not a system file
//!     CanModifyFile = Not ((attr And vbReadOnly) Or (attr And vbSystem))
//!     Exit Function
//!     
//! ErrorHandler:
//!     CanModifyFile = False
//! End Function
//!
//! Sub ModifyFileIfAllowed(filename As String)
//!     If CanModifyFile(filename) Then
//!         ' Proceed with modification
//!         Debug.Print "File can be modified"
//!     Else
//!         Debug.Print "File is protected"
//!     End If
//! End Sub
//! ```
//!
//! ## 9. Directory Tree Walker
//!
//! ```vb
//! Sub WalkDirectory(path As String, Optional level As Integer = 0)
//!     Dim filename As String
//!     Dim fullPath As String
//!     Dim attr As Integer
//!     Dim indent As String
//!     
//!     indent = String(level * 2, " ")
//!     filename = Dir(path & "\*.*", vbDirectory)
//!     
//!     Do While filename <> ""
//!         If filename <> "." And filename <> ".." Then
//!             fullPath = path & "\" & filename
//!             
//!             On Error Resume Next
//!             attr = GetAttr(fullPath)
//!             On Error GoTo 0
//!             
//!             If attr And vbDirectory Then
//!                 Debug.Print indent & "[DIR] " & filename
//!                 WalkDirectory fullPath, level + 1
//!             Else
//!                 Debug.Print indent & filename
//!             End If
//!         End If
//!         
//!         filename = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ## 10. File Attribute Report
//!
//! ```vb
//! Sub GenerateAttributeReport(folderPath As String)
//!     Dim filename As String
//!     Dim fullPath As String
//!     Dim attr As Integer
//!     Dim normalCount As Long
//!     Dim readOnlyCount As Long
//!     Dim hiddenCount As Long
//!     Dim systemCount As Long
//!     Dim archiveCount As Long
//!     
//!     filename = Dir(folderPath & "\*.*", vbNormal + vbHidden + vbSystem)
//!     
//!     Do While filename <> ""
//!         fullPath = folderPath & "\" & filename
//!         
//!         On Error Resume Next
//!         attr = GetAttr(fullPath)
//!         On Error GoTo 0
//!         
//!         If Not (attr And vbDirectory) Then
//!             If attr = vbNormal Then normalCount = normalCount + 1
//!             If attr And vbReadOnly Then readOnlyCount = readOnlyCount + 1
//!             If attr And vbHidden Then hiddenCount = hiddenCount + 1
//!             If attr And vbSystem Then systemCount = systemCount + 1
//!             If attr And vbArchive Then archiveCount = archiveCount + 1
//!         End If
//!         
//!         filename = Dir
//!     Loop
//!     
//!     Debug.Print "File Attribute Report for: " & folderPath
//!     Debug.Print String(50, "=")
//!     Debug.Print "Normal: " & normalCount
//!     Debug.Print "Read-Only: " & readOnlyCount
//!     Debug.Print "Hidden: " & hiddenCount
//!     Debug.Print "System: " & systemCount
//!     Debug.Print "Archive: " & archiveCount
//! End Sub
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. File Attribute Manager Class
//!
//! ```vb
//! ' Class: FileAttributeManager
//! Private m_Filename As String
//! Private m_Attributes As Integer
//!
//! Public Sub LoadFile(filename As String)
//!     m_Filename = filename
//!     RefreshAttributes
//! End Sub
//!
//! Private Sub RefreshAttributes()
//!     On Error Resume Next
//!     m_Attributes = GetAttr(m_Filename)
//!     On Error GoTo 0
//! End Sub
//!
//! Public Property Get IsReadOnly() As Boolean
//!     IsReadOnly = (m_Attributes And vbReadOnly) <> 0
//! End Property
//!
//! Public Property Let IsReadOnly(value As Boolean)
//!     If value Then
//!         SetAttr m_Filename, m_Attributes Or vbReadOnly
//!     Else
//!         SetAttr m_Filename, m_Attributes And Not vbReadOnly
//!     End If
//!     RefreshAttributes
//! End Property
//!
//! Public Property Get IsHidden() As Boolean
//!     IsHidden = (m_Attributes And vbHidden) <> 0
//! End Property
//!
//! Public Property Let IsHidden(value As Boolean)
//!     If value Then
//!         SetAttr m_Filename, m_Attributes Or vbHidden
//!     Else
//!         SetAttr m_Filename, m_Attributes And Not vbHidden
//!     End If
//!     RefreshAttributes
//! End Property
//!
//! Public Property Get IsDirectory() As Boolean
//!     IsDirectory = (m_Attributes And vbDirectory) <> 0
//! End Property
//!
//! Public Property Get AttributeValue() As Integer
//!     AttributeValue = m_Attributes
//! End Property
//! ```
//!
//! ## 2. Smart File Filter
//!
//! ```vb
//! Type FileFilter
//!     IncludeReadOnly As Boolean
//!     IncludeHidden As Boolean
//!     IncludeSystem As Boolean
//!     IncludeArchive As Boolean
//!     ExcludeDirectories As Boolean
//! End Type
//!
//! Function MatchesFilter(filename As String, filter As FileFilter) As Boolean
//!     Dim attr As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr = GetAttr(filename)
//!     
//!     ' Check exclusions first
//!     If filter.ExcludeDirectories And (attr And vbDirectory) Then
//!         MatchesFilter = False
//!         Exit Function
//!     End If
//!     
//!     ' Check inclusions
//!     MatchesFilter = True
//!     
//!     If Not filter.IncludeReadOnly And (attr And vbReadOnly) Then
//!         MatchesFilter = False
//!     ElseIf Not filter.IncludeHidden And (attr And vbHidden) Then
//!         MatchesFilter = False
//!     ElseIf Not filter.IncludeSystem And (attr And vbSystem) Then
//!         MatchesFilter = False
//!     ElseIf Not filter.IncludeArchive And (attr And vbArchive) Then
//!         MatchesFilter = False
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     MatchesFilter = False
//! End Function
//! ```
//!
//! ## 3. Attribute Change Monitor
//!
//! ```vb
//! Type FileSnapshot
//!     Filename As String
//!     Attributes As Integer
//!     Timestamp As Date
//! End Type
//!
//! Private m_Snapshots As Collection
//!
//! Sub InitializeMonitor()
//!     Set m_Snapshots = New Collection
//! End Sub
//!
//! Sub TakeSnapshot(filename As String)
//!     Dim snapshot As FileSnapshot
//!     
//!     snapshot.Filename = filename
//!     snapshot.Attributes = GetAttr(filename)
//!     snapshot.Timestamp = Now
//!     
//!     m_Snapshots.Add snapshot, filename
//! End Sub
//!
//! Function DetectAttributeChanges(filename As String) As Boolean
//!     Dim currentAttr As Integer
//!     Dim snapshot As FileSnapshot
//!     Dim i As Long
//!     
//!     currentAttr = GetAttr(filename)
//!     
//!     ' Find snapshot
//!     For i = 1 To m_Snapshots.Count
//!         snapshot = m_Snapshots(i)
//!         If snapshot.Filename = filename Then
//!             DetectAttributeChanges = (snapshot.Attributes <> currentAttr)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     DetectAttributeChanges = False
//! End Function
//! ```
//!
//! ## 4. Bulk Attribute Operations
//!
//! ```vb
//! Sub SetAttributeForMultipleFiles(filenames() As String, _
//!                                  attributeToSet As Integer, _
//!                                  enable As Boolean)
//!     Dim i As Long
//!     Dim currentAttr As Integer
//!     Dim newAttr As Integer
//!     
//!     For i = LBound(filenames) To UBound(filenames)
//!         On Error Resume Next
//!         
//!         currentAttr = GetAttr(filenames(i))
//!         
//!         If enable Then
//!             newAttr = currentAttr Or attributeToSet
//!         Else
//!             newAttr = currentAttr And Not attributeToSet
//!         End If
//!         
//!         SetAttr filenames(i), newAttr
//!         
//!         On Error GoTo 0
//!     Next i
//! End Sub
//!
//! ' Usage
//! Sub MakeFilesReadOnly()
//!     Dim files() As String
//!     files = Array("file1.txt", "file2.txt", "file3.txt")
//!     
//!     SetAttributeForMultipleFiles files, vbReadOnly, True
//! End Sub
//! ```
//!
//! ## 5. File Security Analyzer
//!
//! ```vb
//! Type SecurityIssue
//!     Filename As String
//!     IssueType As String
//!     Severity As String
//! End Type
//!
//! Function AnalyzeFileSecurity(folderPath As String) As Collection
//!     Dim issues As New Collection
//!     Dim filename As String
//!     Dim fullPath As String
//!     Dim attr As Integer
//!     Dim issue As SecurityIssue
//!     
//!     filename = Dir(folderPath & "\*.*", vbNormal + vbHidden + vbSystem)
//!     
//!     Do While filename <> ""
//!         fullPath = folderPath & "\" & filename
//!         
//!         On Error Resume Next
//!         attr = GetAttr(fullPath)
//!         On Error GoTo 0
//!         
//!         If Not (attr And vbDirectory) Then
//!             ' Check for security issues
//!             
//!             ' Issue 1: System files that are not hidden
//!             If (attr And vbSystem) And Not (attr And vbHidden) Then
//!                 issue.Filename = filename
//!                 issue.IssueType = "System file not hidden"
//!                 issue.Severity = "Medium"
//!                 issues.Add issue
//!             End If
//!             
//!             ' Issue 2: Important files not marked read-only
//!             If InStr(1, filename, "config", vbTextCompare) > 0 Then
//!                 If Not (attr And vbReadOnly) Then
//!                     issue.Filename = filename
//!                     issue.IssueType = "Config file not read-only"
//!                     issue.Severity = "Low"
//!                     issues.Add issue
//!                 End If
//!             End If
//!         End If
//!         
//!         filename = Dir
//!     Loop
//!     
//!     Set AnalyzeFileSecurity = issues
//! End Function
//! ```
//!
//! ## 6. Attribute Comparison Tool
//!
//! ```vb
//! Function CompareFileAttributes(file1 As String, file2 As String) As String
//!     Dim attr1 As Integer
//!     Dim attr2 As Integer
//!     Dim differences As String
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     attr1 = GetAttr(file1)
//!     attr2 = GetAttr(file2)
//!     
//!     differences = ""
//!     
//!     If (attr1 And vbReadOnly) <> (attr2 And vbReadOnly) Then
//!         differences = differences & "Read-Only differs" & vbCrLf
//!     End If
//!     
//!     If (attr1 And vbHidden) <> (attr2 And vbHidden) Then
//!         differences = differences & "Hidden differs" & vbCrLf
//!     End If
//!     
//!     If (attr1 And vbSystem) <> (attr2 And vbSystem) Then
//!         differences = differences & "System differs" & vbCrLf
//!     End If
//!     
//!     If (attr1 And vbArchive) <> (attr2 And vbArchive) Then
//!         differences = differences & "Archive differs" & vbCrLf
//!     End If
//!     
//!     If differences = "" Then
//!         CompareFileAttributes = "Attributes are identical"
//!     Else
//!         CompareFileAttributes = "Differences found:" & vbCrLf & differences
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     CompareFileAttributes = "Error: " & Err.Description
//! End Function
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeGetAttr(pathname As String, _
//!                      Optional defaultValue As Integer = -1) As Integer
//!     On Error GoTo ErrorHandler
//!     
//!     SafeGetAttr = GetAttr(pathname)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 53  ' File not found
//!             Debug.Print "File not found: " & pathname
//!         Case 76  ' Path not found
//!             Debug.Print "Path not found: " & pathname
//!         Case 68  ' Device unavailable
//!             Debug.Print "Device unavailable: " & pathname
//!         Case Else
//!             Debug.Print "Error " & Err.Number & ": " & Err.Description
//!     End Select
//!     
//!     SafeGetAttr = defaultValue
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 53 (File not found)**: The specified file doesn't exist.
//! - **Error 76 (Path not found)**: The specified path doesn't exist.
//! - **Error 68 (Device unavailable)**: Network drive or removable media not available.
//!
//! # Performance Considerations
//!
//! - `GetAttr` is a fast file system call
//! - Use `Dir` to check existence before calling `GetAttr` if unsure
//! - For multiple files, consider caching attribute values
//! - Network paths may be slower than local paths
//! - Accessing removable media can cause delays if not present
//!
//! # Best Practices
//!
//! 1. **Always use error handling** - file may not exist
//! 2. **Use bitwise AND** to test individual attributes
//! 3. **Check Dir first** if file existence is uncertain
//! 4. **Cache results** when checking attributes repeatedly
//! 5. **Use constants** (vbReadOnly, etc.) instead of numeric values
//! 6. **Handle network paths carefully** - may be unavailable
//! 7. **Don't assume attributes** - always check explicitly
//!
//! # Comparison with Other Functions
//!
//! ## `GetAttr` vs `FileLen`
//!
//! ```vb
//! ' GetAttr - Returns file attributes
//! attr = GetAttr("file.txt")  ' Returns attribute flags
//!
//! ' FileLen - Returns file size
//! size = FileLen("file.txt")  ' Returns size in bytes
//! ```
//!
//! ## `GetAttr` vs `Dir`
//!
//! ```vb
//! ' GetAttr - Gets attributes of specific file
//! attr = GetAttr("file.txt")
//!
//! ' Dir - Searches for files matching pattern
//! filename = Dir("*.txt", vbNormal)
//! ```
//!
//! ## `GetAttr` vs `FileDateTime`
//!
//! ```vb
//! ' GetAttr - Returns attributes
//! attr = GetAttr("file.txt")
//!
//! ' FileDateTime - Returns modification date/time
//! dt = FileDateTime("file.txt")
//! ```
//!
//! # Limitations
//!
//! - Returns Integer (limited to values 0-32767 due to VB6 Integer type)
//! - `vbAlias` only works on Macintosh systems
//! - Cannot set attributes (use `SetAttr` for that)
//! - No support for extended attributes or NTFS features
//! - Limited to file system attributes only
//! - Does not provide security descriptor information
//!
//! # Bitwise Operations
//!
//! Testing for specific attributes:
//!
//! ```vb
//! ' Test if read-only
//! If (attr And vbReadOnly) Then ' Has read-only attribute
//!
//! ' Test if NOT read-only
//! If Not (attr And vbReadOnly) Then ' Doesn't have read-only
//!
//! ' Test for multiple attributes
//! If (attr And (vbReadOnly Or vbHidden)) Then ' Has either attribute
//!
//! ' Test for exact attribute combination
//! If attr = (vbReadOnly Or vbHidden) Then ' Has exactly these two
//! ```
//!
//! # Related Functions
//!
//! - `SetAttr` - Sets the attributes of a file
//! - `Dir` - Returns a file or directory name matching a pattern
//! - `FileLen` - Returns the length of a file in bytes
//! - `FileDateTime` - Returns the date and time a file was created or last modified
//! - `FileAttr` - Returns the file mode or file handle for an open file
//! - `FileExists` - Checks if a file exists (custom function using Dir)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_getattr_basic() {
        let source = r#"attr = GetAttr("C:\data.txt")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_with_variable() {
        let source = r#"fileAttr = GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_readonly_check() {
        let source = r#"If GetAttr("file.txt") And vbReadOnly Then MsgBox "Read-only""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_directory_check() {
        let source = r#"If GetAttr(path) And vbDirectory Then isDir = True"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_hidden_check() {
        let source = r#"If GetAttr(fullPath) And vbHidden Then Debug.Print "Hidden""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_in_function() {
        let source = r#"Function IsDirectory(path As String) As Boolean
    IsDirectory = (GetAttr(path) And vbDirectory) <> 0
End Function"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_assignment() {
        let source = r#"Dim attr As Integer
attr = GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_error_handling() {
        let source = r#"On Error GoTo ErrorHandler
attr = GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_system_check() {
        let source = r#"If GetAttr(filename) And vbSystem Then Exit Sub"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_archive_check() {
        let source = r#"needsBackup = (GetAttr(filename) And vbArchive) <> 0"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_multiple_checks() {
        let source = r#"If GetAttr(file) And vbReadOnly Then description = "Read-Only""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_comparison() {
        let source = r#"If GetAttr(filename) = vbNormal Then MsgBox "Normal file""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_not_operator() {
        let source = r#"canModify = Not (GetAttr(filename) And vbReadOnly)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_select_case() {
        let source = r#"Select Case GetAttr(filename) And vbDirectory
    Case 0
        Debug.Print "File"
End Select"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_do_loop() {
        let source = r#"Do While filename <> ""
    attr = GetAttr(fullPath)
Loop"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_setattr() {
        let source = r#"SetAttr filename, GetAttr(filename) And Not vbReadOnly"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_debug_print() {
        let source = r#"Debug.Print "Attributes: " & GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_class_member() {
        let source = r#"m_Attributes = GetAttr(m_Filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_property() {
        let source = r#"IsReadOnly = (GetAttr(filename) And vbReadOnly) <> 0"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_type_field() {
        let source = r#"snapshot.Attributes = GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_or_operator() {
        let source = r#"newAttr = GetAttr(filename) Or vbReadOnly"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_comparison_vars() {
        let source = r#"If GetAttr(file1) <> GetAttr(file2) Then changed = True"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_for_loop() {
        let source = r#"For i = LBound(files) To UBound(files)
    currentAttr = GetAttr(files(i))
Next i"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_msgbox() {
        let source = r#"MsgBox "File attributes: " & GetAttr(filename)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_listbox() {
        let source = r#"lst.AddItem filename & " (" & GetAttr(fullPath) & ")""#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_iif() {
        let source = r#"result = IIf(GetAttr(filename) And vbReadOnly, "RO", "RW")"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_getattr_on_error_resume() {
        let source = r#"On Error Resume Next
attr = GetAttr(fullPath)
On Error GoTo 0"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("GetAttr"));
        assert!(debug.contains("Identifier"));
    }
}
