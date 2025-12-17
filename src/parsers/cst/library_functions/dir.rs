//! # `Dir` Function
//!
//! Returns a `String` representing the name of a file, directory, or folder that matches a
//! specified pattern or file attribute, or the volume label of a drive.
//!
//! ## Syntax
//!
//! ```vb
//! Dir[(pathname[, attributes])]
//! ```
//!
//! ## Parameters
//!
//! - **pathname**: Optional. `String` expression that specifies a file name, directory name,
//!   or folder name. May include wildcards (* and ?). If not specified, uses the pattern
//!   from the previous `Dir` call.
//! - **attributes**: Optional. Constant or numeric expression whose sum specifies file
//!   attributes. If omitted, returns files that match pathname but have no attributes.
//!
//! ## Attributes
//!
//! - **vbNormal** (0): Normal files (default)
//! - **vbReadOnly** (1): Read-only files
//! - **vbHidden** (2): Hidden files
//! - **vbSystem** (4): System files
//! - **vbVolume** (8): Volume label (pathname ignored)
//! - **vbDirectory** (16): Directories or folders
//! - **vbArchive** (32): Files that have changed since last backup
//!
//! ## Return Value
//!
//! Returns a `String` containing the name of a file, directory, or folder that matches the
//! specified pattern and attributes. Returns a zero-length string ("") when no more files
//! are found.
//!
//! ## Remarks
//!
//! The `Dir` function is used to retrieve file and directory names that match a pattern.
//! It's commonly used to iterate through files in a directory or to check if a file exists.
//!
//! **Important Characteristics:**
//!
//! - First call with pathname initializes search and returns first match
//! - Subsequent calls without arguments return next matching file
//! - Returns empty string ("") when no more matches found
//! - Case-insensitive pattern matching
//! - Supports wildcards: * (multiple chars) and ? (single char)
//! - Does not return "." and ".." directory entries
//! - Order of returned files is not guaranteed (typically file system order)
//! - Maintains internal state between calls
//! - Multiple Dir loops cannot be nested without complications
//! - Changing directory during Dir enumeration can cause issues
//!
//! ## Wildcards
//!
//! - `*` - Matches zero or more characters
//! - `?` - Matches exactly one character
//! - `*.*` - All files with extensions
//! - `*.txt` - All .txt files
//! - `test?.dat` - Files like test1.dat, testA.dat
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get first .txt file in current directory
//! Dim fileName As String
//! fileName = Dir("*.txt")
//! MsgBox fileName
//!
//! ' Check if specific file exists
//! If Len(Dir("C:\data\report.txt")) > 0 Then
//!     MsgBox "File exists"
//! End If
//!
//! ' Get volume label
//! Dim volumeLabel As String
//! volumeLabel = Dir("C:\", vbVolume)
//! ```
//!
//! ### Iterate Through Files
//!
//! ```vb
//! Sub ListAllTextFiles()
//!     Dim fileName As String
//!     
//!     ' First call with pattern
//!     fileName = Dir("C:\Documents\*.txt")
//!     
//!     ' Loop through all matches
//!     Do While fileName <> ""
//!         Debug.Print fileName
//!         fileName = Dir  ' Subsequent calls without arguments
//!     Loop
//! End Sub
//! ```
//!
//! ### Count Files
//!
//! ```vb
//! Function CountFiles(path As String, pattern As String) As Long
//!     Dim fileName As String
//!     Dim count As Long
//!     
//!     count = 0
//!     fileName = Dir(path & "\" & pattern)
//!     
//!     Do While fileName <> ""
//!         count = count + 1
//!         fileName = Dir
//!     Loop
//!     
//!     CountFiles = count
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### File Existence Check
//!
//! ```vb
//! Function FileExists(filePath As String) As Boolean
//!     FileExists = (Len(Dir(filePath)) > 0)
//! End Function
//!
//! ' Usage
//! If FileExists("C:\data\file.txt") Then
//!     ' File exists
//! End If
//! ```
//!
//! ### Get All Files in Directory
//!
//! ```vb
//! Function GetFileList(folderPath As String, pattern As String) As Variant
//!     Dim files() As String
//!     Dim fileName As String
//!     Dim count As Long
//!     
//!     count = 0
//!     ReDim files(0 To 100)
//!     
//!     fileName = Dir(folderPath & "\" & pattern)
//!     Do While fileName <> ""
//!         files(count) = fileName
//!         count = count + 1
//!         
//!         If count > UBound(files) Then
//!             ReDim Preserve files(0 To UBound(files) + 100)
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve files(0 To count - 1)
//!         GetFileList = files
//!     Else
//!         GetFileList = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Find Files by Attribute
//!
//! ```vb
//! Sub ListHiddenFiles(folderPath As String)
//!     Dim fileName As String
//!     
//!     fileName = Dir(folderPath & "\*.*", vbHidden)
//!     Do While fileName <> ""
//!         Debug.Print "Hidden: " & fileName
//!         fileName = Dir
//!     Loop
//! End Sub
//!
//! Sub ListDirectories(folderPath As String)
//!     Dim dirName As String
//!     
//!     dirName = Dir(folderPath & "\*.*", vbDirectory)
//!     Do While dirName <> ""
//!         ' Filter out "." and ".." if they appear
//!         If dirName <> "." And dirName <> ".." Then
//!             ' Check if it's actually a directory
//!             If GetAttr(folderPath & "\" & dirName) And vbDirectory Then
//!                 Debug.Print "Directory: " & dirName
//!             End If
//!         End If
//!         dirName = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ### Search Multiple File Types
//!
//! ```vb
//! Function FindDocuments(folderPath As String) As Variant
//!     Dim files() As String
//!     Dim fileName As String
//!     Dim count As Long
//!     Dim extensions As Variant
//!     Dim i As Integer
//!     
//!     extensions = Array("*.txt", "*.doc", "*.docx", "*.pdf")
//!     ReDim files(0 To 100)
//!     count = 0
//!     
//!     For i = LBound(extensions) To UBound(extensions)
//!         fileName = Dir(folderPath & "\" & extensions(i))
//!         Do While fileName <> ""
//!             files(count) = fileName
//!             count = count + 1
//!             
//!             If count > UBound(files) Then
//!                 ReDim Preserve files(0 To UBound(files) + 100)
//!             End If
//!             
//!             fileName = Dir
//!         Loop
//!     Next i
//!     
//!     If count > 0 Then
//!         ReDim Preserve files(0 To count - 1)
//!         FindDocuments = files
//!     Else
//!         FindDocuments = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Get Full File Paths
//!
//! ```vb
//! Function GetFullPaths(folderPath As String, pattern As String) As Variant
//!     Dim paths() As String
//!     Dim fileName As String
//!     Dim count As Long
//!     
//!     count = 0
//!     ReDim paths(0 To 100)
//!     
//!     fileName = Dir(folderPath & "\" & pattern)
//!     Do While fileName <> ""
//!         paths(count) = folderPath & "\" & fileName
//!         count = count + 1
//!         
//!         If count > UBound(paths) Then
//!             ReDim Preserve paths(0 To UBound(paths) + 100)
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve paths(0 To count - 1)
//!         GetFullPaths = paths
//!     Else
//!         GetFullPaths = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Delete All Files Matching Pattern
//!
//! ```vb
//! Sub DeleteMatchingFiles(folderPath As String, pattern As String)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     
//!     fileName = Dir(folderPath & "\" & pattern)
//!     Do While fileName <> ""
//!         fullPath = folderPath & "\" & fileName
//!         
//!         ' Get next file BEFORE deleting (Dir state would be lost)
//!         fileName = Dir
//!         
//!         ' Delete the file
//!         Kill fullPath
//!     Loop
//! End Sub
//! ```
//!
//! ### Find Newest File
//!
//! ```vb
//! Function GetNewestFile(folderPath As String, pattern As String) As String
//!     Dim fileName As String
//!     Dim newestFile As String
//!     Dim newestDate As Date
//!     Dim currentDate As Date
//!     Dim fullPath As String
//!     
//!     newestDate = 0
//!     fileName = Dir(folderPath & "\" & pattern)
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & "\" & fileName
//!         currentDate = FileDateTime(fullPath)
//!         
//!         If currentDate > newestDate Then
//!             newestDate = currentDate
//!             newestFile = fileName
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     GetNewestFile = newestFile
//! End Function
//! ```
//!
//! ### Calculate Total Size
//!
//! ```vb
//! Function GetTotalFileSize(folderPath As String, pattern As String) As Double
//!     Dim fileName As String
//!     Dim totalSize As Double
//!     Dim fullPath As String
//!     
//!     totalSize = 0
//!     fileName = Dir(folderPath & "\" & pattern)
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & "\" & fileName
//!         totalSize = totalSize + FileLen(fullPath)
//!         fileName = Dir
//!     Loop
//!     
//!     GetTotalFileSize = totalSize
//! End Function
//! ```
//!
//! ### Recursive Directory Search
//!
//! ```vb
//! Sub SearchRecursive(folderPath As String, pattern As String)
//!     Dim fileName As String
//!     Dim dirName As String
//!     Dim fullPath As String
//!     
//!     ' Search files in current directory
//!     fileName = Dir(folderPath & "\" & pattern)
//!     Do While fileName <> ""
//!         Debug.Print folderPath & "\" & fileName
//!         fileName = Dir
//!     Loop
//!     
//!     ' Search subdirectories
//!     dirName = Dir(folderPath & "\*.*", vbDirectory)
//!     Do While dirName <> ""
//!         If dirName <> "." And dirName <> ".." Then
//!             fullPath = folderPath & "\" & dirName
//!             If GetAttr(fullPath) And vbDirectory Then
//!                 SearchRecursive fullPath, pattern
//!             End If
//!         End If
//!         dirName = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### File Filter with Multiple Criteria
//!
//! ```vb
//! Function FindFilesAdvanced(folderPath As String, _
//!                           minSize As Long, maxSize As Long, _
//!                           afterDate As Date) As Variant
//!     Dim files() As String
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     Dim fileDate As Date
//!     Dim count As Long
//!     
//!     count = 0
//!     ReDim files(0 To 100)
//!     
//!     fileName = Dir(folderPath & "\*.*")
//!     Do While fileName <> ""
//!         fullPath = folderPath & "\" & fileName
//!         fileSize = FileLen(fullPath)
//!         fileDate = FileDateTime(fullPath)
//!         
//!         If fileSize >= minSize And fileSize <= maxSize And fileDate > afterDate Then
//!             files(count) = fileName
//!             count = count + 1
//!             
//!             If count > UBound(files) Then
//!                 ReDim Preserve files(0 To UBound(files) + 100)
//!             End If
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve files(0 To count - 1)
//!         FindFilesAdvanced = files
//!     Else
//!         FindFilesAdvanced = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Safe Dir Loop Helper
//!
//! ```vb
//! ' Helper to avoid nested Dir issues
//! Type FileInfo
//!     Name As String
//!     FullPath As String
//!     Size As Long
//!     Modified As Date
//! End Type
//!
//! Function GetFileInfoList(folderPath As String, pattern As String) As Variant
//!     Dim files() As FileInfo
//!     Dim fileName As String
//!     Dim count As Long
//!     
//!     count = 0
//!     ReDim files(0 To 100)
//!     
//!     fileName = Dir(folderPath & "\" & pattern)
//!     Do While fileName <> ""
//!         files(count).Name = fileName
//!         files(count).FullPath = folderPath & "\" & fileName
//!         files(count).Size = FileLen(files(count).FullPath)
//!         files(count).Modified = FileDateTime(files(count).FullPath)
//!         
//!         count = count + 1
//!         If count > UBound(files) Then
//!             ReDim Preserve files(0 To UBound(files) + 100)
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve files(0 To count - 1)
//!         GetFileInfoList = files
//!     Else
//!         GetFileInfoList = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Backup Old Files
//!
//! ```vb
//! Sub BackupOldFiles(sourcePath As String, backupPath As String, daysOld As Integer)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim cutoffDate As Date
//!     
//!     cutoffDate = Date - daysOld
//!     fileName = Dir(sourcePath & "\*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = sourcePath & "\" & fileName
//!         
//!         If FileDateTime(fullPath) < cutoffDate Then
//!             FileCopy fullPath, backupPath & "\" & fileName
//!         End If
//!         
//!         fileName = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ### Build File Index
//!
//! ```vb
//! Function BuildFileIndex(rootPath As String) As Collection
//!     Dim index As New Collection
//!     Dim fileName As String
//!     
//!     ' Add all files to collection with full path as key
//!     fileName = Dir(rootPath & "\*.*")
//!     Do While fileName <> ""
//!         On Error Resume Next
//!         index.Add fileName, UCase(fileName)
//!         On Error GoTo 0
//!         fileName = Dir
//!     Loop
//!     
//!     Set BuildFileIndex = index
//! End Function
//! ```
//!
//! ### File Synchronization Check
//!
//! ```vb
//! Function CompareDirectories(path1 As String, path2 As String) As String
//!     Dim files1 As Collection
//!     Dim files2 As Collection
//!     Dim fileName As String
//!     Dim report As String
//!     
//!     Set files1 = New Collection
//!     Set files2 = New Collection
//!     
//!     ' Get files from first directory
//!     fileName = Dir(path1 & "\*.*")
//!     Do While fileName <> ""
//!         files1.Add fileName
//!         fileName = Dir
//!     Loop
//!     
//!     ' Get files from second directory
//!     fileName = Dir(path2 & "\*.*")
//!     Do While fileName <> ""
//!         files2.Add fileName
//!         fileName = Dir
//!     Loop
//!     
//!     ' Compare (simplified - full version would check both ways)
//!     report = "Files only in " & path1 & ":" & vbCrLf
//!     ' ... comparison logic ...
//!     
//!     CompareDirectories = report
//! End Function
//! ```
//!
//! ### Generate Directory Listing Report
//!
//! ```vb
//! Sub ExportDirectoryListing(folderPath As String, outputFile As String)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileNum As Integer
//!     
//!     fileNum = FreeFile
//!     Open outputFile For Output As #fileNum
//!     
//!     Print #fileNum, "Directory Listing for: " & folderPath
//!     Print #fileNum, "Generated: " & Now
//!     Print #fileNum, String(80, "-")
//!     Print #fileNum, "Filename" & vbTab & "Size" & vbTab & "Modified"
//!     Print #fileNum, String(80, "-")
//!     
//!     fileName = Dir(folderPath & "\*.*")
//!     Do While fileName <> ""
//!         fullPath = folderPath & "\" & fileName
//!         Print #fileNum, fileName & vbTab & _
//!                        FileLen(fullPath) & vbTab & _
//!                        FileDateTime(fullPath)
//!         fileName = Dir
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDir(pathname As String, Optional attributes As Integer = 0) As String
//!     On Error Resume Next
//!     SafeDir = Dir(pathname, attributes)
//!     If Err.Number <> 0 Then
//!         SafeDir = ""
//!     End If
//! End Function
//!
//! Function SafeFileExists(filePath As String) As Boolean
//!     On Error Resume Next
//!     SafeFileExists = (Len(Dir(filePath)) > 0)
//!     If Err.Number <> 0 Then
//!         SafeFileExists = False
//!     End If
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 52** (Bad file name or number): Invalid pathname or pattern
//! - **Error 76** (Path not found): Directory does not exist
//! - **Error 68** (Device unavailable): Drive not ready or network path unavailable
//!
//! ## Performance Considerations
//!
//! - `Dir` is relatively fast for simple file enumeration
//! - For large directories, consider showing progress
//! - Network paths can be slow; consider timeout handling
//! - Avoid nested `Dir` loops (collect to array first)
//! - `FileSystemObject` may be faster for complex operations
//! - Cache results if scanning same directory repeatedly
//!
//! ## Best Practices
//!
//! ### Always Check for Empty String
//!
//! ```vb
//! ' Good - Check for no more files
//! fileName = Dir("*.txt")
//! Do While fileName <> ""
//!     ' Process file
//!     fileName = Dir
//! Loop
//!
//! ' Avoid - May cause infinite loop
//! fileName = Dir("*.txt")
//! Do While Len(fileName) > 0  ' Less reliable
//!     fileName = Dir
//! Loop
//! ```
//!
//! ### Store Files Before Processing
//!
//! ```vb
//! ' Good - Collect files first
//! Dim files() As String, i As Integer
//! ReDim files(0 To 100)
//! count = 0
//!
//! fileName = Dir("*.txt")
//! Do While fileName <> ""
//!     files(count) = fileName
//!     count = count + 1
//!     fileName = Dir
//! Loop
//!
//! ' Now process without Dir active
//! For i = 0 To count - 1
//!     ProcessFile files(i)
//! Next i
//! ```
//!
//! ### Use Absolute Paths
//!
//! ```vb
//! ' Good - Explicit path
//! fileName = Dir("C:\Data\*.txt")
//!
//! ' Risky - Depends on current directory
//! fileName = Dir("*.txt")
//! ```
//!
//! ### Handle No Matches Gracefully
//!
//! ```vb
//! fileName = Dir("*.xyz")
//! If fileName = "" Then
//!     MsgBox "No matching files found"
//!     Exit Sub
//! End If
//! ```
//!
//! ## Comparison with Other Methods
//!
//! ### `Dir` vs `FileSystemObject`
//!
//! ```vb
//! ' Dir - Built-in, faster for simple cases
//! fileName = Dir("C:\Data\*.txt")
//!
//! ' FileSystemObject - More features, but requires reference
//! Dim fso As New FileSystemObject
//! Dim folder As Folder
//! Set folder = fso.GetFolder("C:\Data")
//! ' ... more complex but more powerful
//! ```
//!
//! ### `Dir` vs `File Dialog`
//!
//! ```vb
//! ' Dir - Programmatic file discovery
//! fileName = Dir("*.txt")
//!
//! ' File Dialog - User selection
//! fileName = Application.GetOpenFilename("Text Files (*.txt), *.txt")
//! ```
//!
//! ## Limitations
//!
//! - Cannot nest Dir loops reliably (single internal state)
//! - Does not return files in sorted order
//! - Returns only file names, not full paths
//! - No built-in recursion into subdirectories
//! - Cannot filter by date, size, or other attributes directly
//! - Changing current directory during enumeration causes issues
//! - Limited attribute filtering compared to `FileSystemObject`
//!
//! ## Related Functions
//!
//! - `GetAttr`: Gets file attributes
//! - `SetAttr`: Sets file attributes
//! - `FileLen`: Returns file size
//! - `FileDateTime`: Returns file modification date/time
//! - `CurDir`: Returns current directory
//! - `ChDir`: Changes current directory
//! - `MkDir`: Creates directory
//! - `RmDir`: Removes directory
//! - `Kill`: Deletes file
//! - `FileCopy`: Copies file
//! - `Name`: Renames/moves file

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn dir_basic() {
        let source = r#"
fileName = Dir("*.txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_path() {
        let source = r#"
fileName = Dir("C:\Data\*.txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_attributes() {
        let source = r#"
fileName = Dir("*.*", vbDirectory)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_no_arguments() {
        let source = r#"
fileName = Dir
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_in_loop() {
        let source = r#"
fileName = Dir("*.txt")
Do While fileName <> ""
    Debug.Print fileName
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_file_exists() {
        let source = r#"
If Len(Dir("C:\file.txt")) > 0 Then
    MsgBox "File exists"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_volume_label() {
        let source = r#"
volumeLabel = Dir("C:\", vbVolume)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_in_function() {
        let source = r#"
Function FileExists(filePath As String) As Boolean
    FileExists = (Len(Dir(filePath)) > 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_hidden_files() {
        let source = r#"
fileName = Dir("*.*", vbHidden)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_multiple_attributes() {
        let source = r#"
fileName = Dir("*.*", vbHidden + vbSystem)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_variable() {
        let source = r#"
pattern = "*.doc"
fileName = Dir(pattern)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_concatenated_path() {
        let source = r#"
fileName = Dir(folderPath & "\*.txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_count_files() {
        let source = r#"
count = 0
fileName = Dir("*.txt")
Do While fileName <> ""
    count = count + 1
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_array_population() {
        let source = r#"
fileName = Dir("*.txt")
Do While fileName <> ""
    files(count) = fileName
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_kill() {
        let source = r#"
fileName = Dir("*.tmp")
Do While fileName <> ""
    Kill folderPath & "\" & fileName
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_filelen() {
        let source = r#"
fileName = Dir("*.*")
Do While fileName <> ""
    totalSize = totalSize + FileLen(fileName)
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_error_handling() {
        let source = r#"
On Error Resume Next
fileName = Dir("C:\NonExistent\*.*")
If Err.Number <> 0 Then
    MsgBox "Error accessing directory"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_select_case() {
        let source = r#"
Select Case UCase(Right(Dir("*.*"), 3))
    Case "TXT"
        MsgBox "Text file"
    Case "DOC"
        MsgBox "Word file"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_getattr() {
        let source = r#"
fileName = Dir("*.*", vbDirectory)
Do While fileName <> ""
    If GetAttr(fileName) And vbDirectory Then
        Debug.Print fileName
    End If
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_msgbox() {
        let source = r#"
MsgBox "First file: " & Dir("*.txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_parentheses_optional() {
        let source = r#"
fileName = Dir()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_nested_paths() {
        let source = r#"
fileName = Dir("C:\Users\Documents\Reports\*.pdf")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_wildcard_question() {
        let source = r#"
fileName = Dir("test?.txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_all_files() {
        let source = r#"
fileName = Dir("*.*")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dir_with_filedatetime() {
        let source = r#"
fileName = Dir("*.*")
Do While fileName <> ""
    If FileDateTime(fileName) > cutoffDate Then
        ProcessFile fileName
    End If
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Dir"));
        assert!(debug.contains("Identifier"));
    }
}
