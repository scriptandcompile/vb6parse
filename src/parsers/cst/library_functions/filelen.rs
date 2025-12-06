//! # `FileLen` Function
//!
//! Returns a `Long` specifying the length of a file in bytes.
//!
//! ## Syntax
//!
//! ```vb
//! FileLen(pathname)
//! ```
//!
//! ## Parameters
//!
//! - **pathname**: Required. A `String` expression that specifies a file name. May include
//!   directory or folder, and drive. If the file is not found, an error occurs.
//!
//! ## Return Value
//!
//! Returns a `Long` representing the length of the file in bytes. For open files, the value
//! returned is the size of the file immediately before it was opened.
//!
//! ## Remarks
//!
//! The `FileLen` function returns the size of a file in bytes. This is useful for
//! determining file sizes before reading, checking disk space requirements, validating
//! file downloads, and managing storage.
//!
//! **Important Characteristics:**
//!
//! - Returns file size in bytes
//! - File does not need to be open
//! - Error if file does not exist (Error 53)
//! - Error if path is invalid (Error 76)
//! - Works with full paths and relative paths
//! - For open files, returns size before opening
//! - Maximum file size: 2,147,483,647 bytes (2GB - 1) due to `Long` limit
//! - Returns 0 for empty files
//! - Does not include file system overhead
//! - Can be used with wildcards via `Dir` function
//!
//! ## Typical Uses
//!
//! - Check available space before file operations
//! - Validate file downloads (compare expected vs actual size)
//! - Display file sizes to users
//! - Filter files by size
//! - Calculate total directory size
//! - Determine buffer sizes for file reading
//! - Progress bar calculations for file operations
//! - Detect truncated or corrupted files
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim fileSize As Long
//!
//! ' Get file size in bytes
//! fileSize = FileLen("C:\data.txt")
//! Debug.Print "File size: " & fileSize & " bytes"
//!
//! ' Convert to KB, MB, GB
//! Debug.Print "Size in KB: " & Format(fileSize / 1024, "0.00")
//! Debug.Print "Size in MB: " & Format(fileSize / 1048576, "0.00")
//! ```
//!
//! ### Format File Size for Display
//!
//! ```vb
//! Function FormatFileSize(bytes As Long) As String
//!     Const KB = 1024
//!     Const MB = 1048576         ' 1024 * 1024
//!     Const GB = 1073741824      ' 1024 * 1024 * 1024
//!     
//!     If bytes >= GB Then
//!         FormatFileSize = Format(bytes / GB, "0.00") & " GB"
//!     ElseIf bytes >= MB Then
//!         FormatFileSize = Format(bytes / MB, "0.00") & " MB"
//!     ElseIf bytes >= KB Then
//!         FormatFileSize = Format(bytes / KB, "0.00") & " KB"
//!     Else
//!         FormatFileSize = bytes & " bytes"
//!     End If
//! End Function
//!
//! ' Usage
//! Debug.Print FormatFileSize(FileLen("C:\data.txt"))
//! ```
//!
//! ### Check If File Exists and Get Size
//!
//! ```vb
//! Function GetFileSize(filePath As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     GetFileSize = FileLen(filePath)
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileSize = -1  ' Indicate error
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Calculate Directory Size
//!
//! ```vb
//! Function GetDirectorySize(folderPath As String) As Long
//!     Dim fileName As String
//!     Dim totalSize As Long
//!     Dim fileSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     totalSize = 0
//!     
//!     Do While fileName <> ""
//!         On Error Resume Next
//!         fileSize = FileLen(folderPath & fileName)
//!         
//!         If Err.Number = 0 Then
//!             totalSize = totalSize + fileSize
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     GetDirectorySize = totalSize
//! End Function
//! ```
//!
//! ### Find Largest File
//!
//! ```vb
//! Function FindLargestFile(folderPath As String) As String
//!     Dim fileName As String
//!     Dim largestFile As String
//!     Dim largestSize As Long
//!     Dim currentSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     largestSize = 0
//!     
//!     Do While fileName <> ""
//!         On Error Resume Next
//!         currentSize = FileLen(folderPath & fileName)
//!         
//!         If Err.Number = 0 And currentSize > largestSize Then
//!             largestSize = currentSize
//!             largestFile = fileName
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     FindLargestFile = largestFile
//! End Function
//! ```
//!
//! ### Filter Files by Size
//!
//! ```vb
//! Function GetFilesBySize(folderPath As String, minSize As Long, _
//!                         maxSize As Long) As Collection
//!     Dim files As New Collection
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             If fileSize >= minSize And fileSize <= maxSize Then
//!                 files.Add fullPath
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     Set GetFilesBySize = files
//! End Function
//! ```
//!
//! ### Validate File Download
//!
//! ```vb
//! Function ValidateDownload(filePath As String, expectedSize As Long) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim actualSize As Long
//!     actualSize = FileLen(filePath)
//!     
//!     ValidateDownload = (actualSize = expectedSize)
//!     
//!     If Not ValidateDownload Then
//!         Debug.Print "Size mismatch - Expected: " & expectedSize & _
//!                     ", Actual: " & actualSize
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     ValidateDownload = False
//! End Function
//! ```
//!
//! ### Check Available Space
//!
//! ```vb
//! Function HasEnoughSpace(filePath As String, drive As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileSize As Long
//!     Dim freeSpace As Currency
//!     Dim fso As Object
//!     
//!     ' Get file size
//!     fileSize = FileLen(filePath)
//!     
//!     ' Get free space (using FileSystemObject)
//!     Set fso = CreateObject("Scripting.FileSystemObject")
//!     freeSpace = fso.GetDrive(drive).FreeSpace
//!     
//!     HasEnoughSpace = (freeSpace > fileSize)
//!     Exit Function
//!     
//! ErrorHandler:
//!     HasEnoughSpace = False
//! End Function
//! ```
//!
//! ### Progress Bar for File Copy
//!
//! ```vb
//! Sub CopyFileWithProgress(sourceFile As String, destFile As String)
//!     Dim fileSize As Long
//!     Dim buffer() As Byte
//!     Dim sourceNum As Integer
//!     Dim destNum As Integer
//!     Dim bytesRead As Long
//!     Dim chunkSize As Long
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Get total file size
//!     fileSize = FileLen(sourceFile)
//!     
//!     If fileSize = 0 Then Exit Sub
//!     
//!     chunkSize = 65536  ' 64KB chunks
//!     bytesRead = 0
//!     
//!     sourceNum = FreeFile
//!     Open sourceFile For Binary As #sourceNum
//!     
//!     destNum = FreeFile
//!     Open destFile For Binary As #destNum
//!     
//!     Do While bytesRead < fileSize
//!         If fileSize - bytesRead < chunkSize Then
//!             chunkSize = fileSize - bytesRead
//!         End If
//!         
//!         ReDim buffer(0 To chunkSize - 1)
//!         Get #sourceNum, , buffer
//!         Put #destNum, , buffer
//!         
//!         bytesRead = bytesRead + chunkSize
//!         
//!         ' Update progress bar
//!         ProgressBar.Value = (bytesRead / fileSize) * 100
//!         DoEvents
//!     Loop
//!     
//!     Close #sourceNum
//!     Close #destNum
//!     Exit Sub
//!     
//! ErrorHandler:
//!     If sourceNum > 0 Then Close #sourceNum
//!     If destNum > 0 Then Close #destNum
//! End Sub
//! ```
//!
//! ### List Files with Sizes
//!
//! ```vb
//! Sub ListFilesWithSizes(folderPath As String)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     Debug.Print "Files in: " & folderPath
//!     Debug.Print String(60, "-")
//!     Debug.Print "Filename", "Size"
//!     Debug.Print String(60, "-")
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             Debug.Print fileName, FormatFileSize(fileSize)
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ### Allocate Buffer Based on File Size
//!
//! ```vb
//! Function ReadFileToBuffer(filePath As String) As Byte()
//!     Dim fileSize As Long
//!     Dim buffer() As Byte
//!     Dim fileNum As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Get file size to allocate exact buffer
//!     fileSize = FileLen(filePath)
//!     
//!     If fileSize = 0 Then
//!         ReadFileToBuffer = buffer
//!         Exit Function
//!     End If
//!     
//!     ' Allocate buffer
//!     ReDim buffer(0 To fileSize - 1)
//!     
//!     ' Read file
//!     fileNum = FreeFile
//!     Open filePath For Binary As #fileNum
//!     Get #fileNum, , buffer
//!     Close #fileNum
//!     
//!     ReadFileToBuffer = buffer
//!     Exit Function
//!     
//! ErrorHandler:
//!     If fileNum > 0 Then Close #fileNum
//!     ReadFileToBuffer = buffer
//! End Function
//! ```
//!
//! ### Find Empty Files
//!
//! ```vb
//! Function FindEmptyFiles(folderPath As String) As Collection
//!     Dim emptyFiles As New Collection
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 And fileSize = 0 Then
//!             emptyFiles.Add fullPath
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     Set FindEmptyFiles = emptyFiles
//! End Function
//! ```
//!
//! ### Compare File Sizes
//!
//! ```vb
//! Function CompareFileSizes(file1 As String, file2 As String) As Long
//!     ' Returns: -1 if file1 < file2, 0 if equal, 1 if file1 > file2
//!     On Error GoTo ErrorHandler
//!     
//!     Dim size1 As Long
//!     Dim size2 As Long
//!     
//!     size1 = FileLen(file1)
//!     size2 = FileLen(file2)
//!     
//!     If size1 < size2 Then
//!         CompareFileSizes = -1
//!     ElseIf size1 > size2 Then
//!         CompareFileSizes = 1
//!     Else
//!         CompareFileSizes = 0
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     CompareFileSizes = 0
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Disk Usage Report
//!
//! ```vb
//! Type DirectoryStats
//!     Path As String
//!     FileCount As Long
//!     TotalSize As Long
//!     LargestFile As String
//!     LargestSize As Long
//!     SmallestFile As String
//!     SmallestSize As Long
//!     AverageSize As Long
//! End Type
//!
//! Function AnalyzeDirectory(folderPath As String) As DirectoryStats
//!     Dim stats As DirectoryStats
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     stats.Path = folderPath
//!     stats.FileCount = 0
//!     stats.TotalSize = 0
//!     stats.LargestSize = 0
//!     stats.SmallestSize = 2147483647  ' Max Long
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             stats.FileCount = stats.FileCount + 1
//!             stats.TotalSize = stats.TotalSize + fileSize
//!             
//!             If fileSize > stats.LargestSize Then
//!                 stats.LargestSize = fileSize
//!                 stats.LargestFile = fileName
//!             End If
//!             
//!             If fileSize < stats.SmallestSize Then
//!                 stats.SmallestSize = fileSize
//!                 stats.SmallestFile = fileName
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     If stats.FileCount > 0 Then
//!         stats.AverageSize = stats.TotalSize / stats.FileCount
//!     End If
//!     
//!     AnalyzeDirectory = stats
//! End Function
//! ```
//!
//! ### Sort Files by Size
//!
//! ```vb
//! Type FileSizeInfo
//!     Name As String
//!     Size As Long
//! End Type
//!
//! Function GetFilesSortedBySize(folderPath As String) As Variant
//!     Dim files() As FileSizeInfo
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim count As Long
//!     Dim i As Long, j As Long
//!     Dim temp As FileSizeInfo
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     ReDim files(0 To 100)
//!     count = 0
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         files(count).Name = fileName
//!         files(count).Size = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             count = count + 1
//!             If count > UBound(files) Then
//!                 ReDim Preserve files(0 To UBound(files) + 100)
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve files(0 To count - 1)
//!         
//!         ' Bubble sort by size (largest first)
//!         For i = 0 To count - 2
//!             For j = i + 1 To count - 1
//!                 If files(j).Size > files(i).Size Then
//!                     temp = files(i)
//!                     files(i) = files(j)
//!                     files(j) = temp
//!                 End If
//!             Next j
//!         Next i
//!     End If
//!     
//!     GetFilesSortedBySize = files
//! End Function
//! ```
//!
//! ### File Size Distribution
//!
//! ```vb
//! Type SizeDistribution
//!     Range As String
//!     Count As Long
//!     TotalSize As Long
//! End Type
//!
//! Function GetSizeDistribution(folderPath As String) As Variant
//!     Dim dist(0 To 5) As SizeDistribution
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     ' Define ranges
//!     dist(0).Range = "< 1 KB"
//!     dist(1).Range = "1 KB - 100 KB"
//!     dist(2).Range = "100 KB - 1 MB"
//!     dist(3).Range = "1 MB - 10 MB"
//!     dist(4).Range = "10 MB - 100 MB"
//!     dist(5).Range = "> 100 MB"
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             If fileSize < 1024 Then
//!                 dist(0).Count = dist(0).Count + 1
//!                 dist(0).TotalSize = dist(0).TotalSize + fileSize
//!             ElseIf fileSize < 102400 Then
//!                 dist(1).Count = dist(1).Count + 1
//!                 dist(1).TotalSize = dist(1).TotalSize + fileSize
//!             ElseIf fileSize < 1048576 Then
//!                 dist(2).Count = dist(2).Count + 1
//!                 dist(2).TotalSize = dist(2).TotalSize + fileSize
//!             ElseIf fileSize < 10485760 Then
//!                 dist(3).Count = dist(3).Count + 1
//!                 dist(3).TotalSize = dist(3).TotalSize + fileSize
//!             ElseIf fileSize < 104857600 Then
//!                 dist(4).Count = dist(4).Count + 1
//!                 dist(4).TotalSize = dist(4).TotalSize + fileSize
//!             Else
//!                 dist(5).Count = dist(5).Count + 1
//!                 dist(5).TotalSize = dist(5).TotalSize + fileSize
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     GetSizeDistribution = dist
//! End Function
//! ```
//!
//! ### Quota Management
//!
//! ```vb
//! Function CheckQuota(userFolder As String, quotaLimit As Long) As Boolean
//!     Dim totalSize As Long
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     
//!     If Right(userFolder, 1) <> "\" Then
//!         userFolder = userFolder & "\"
//!     End If
//!     
//!     totalSize = 0
//!     fileName = Dir(userFolder & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = userFolder & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             totalSize = totalSize + fileSize
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     If totalSize > quotaLimit Then
//!         MsgBox "Quota exceeded!" & vbCrLf & _
//!                "Used: " & FormatFileSize(totalSize) & vbCrLf & _
//!                "Limit: " & FormatFileSize(quotaLimit), vbExclamation
//!         CheckQuota = False
//!     Else
//!         CheckQuota = True
//!     End If
//! End Function
//! ```
//!
//! ### Duplicate File Finder (by size)
//!
//! ```vb
//! Function FindPotentialDuplicates(folderPath As String) As Collection
//!     ' Files with same size are potential duplicates
//!     Dim sizeMap As New Collection
//!     Dim duplicates As New Collection
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileSize As Long
//!     Dim sizeKey As String
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileSize = FileLen(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             sizeKey = CStr(fileSize)
//!             
//!             ' Try to add to collection
//!             Err.Clear
//!             sizeMap.Add fullPath, sizeKey
//!             
//!             If Err.Number <> 0 Then
//!                 ' Duplicate size found
//!                 duplicates.Add fullPath
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     Set FindPotentialDuplicates = duplicates
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFileLen(filePath As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     SafeFileLen = FileLen(filePath)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 53  ' File not found
//!             Debug.Print "File not found: " & filePath
//!             SafeFileLen = -1
//!         Case 76  ' Path not found
//!             Debug.Print "Path not found: " & filePath
//!             SafeFileLen = -1
//!         Case Else
//!             Debug.Print "Error " & Err.Number & ": " & Err.Description
//!             SafeFileLen = -1
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 53** (File not found): The specified file does not exist
//! - **Error 76** (Path not found): The specified path is invalid
//! - **Error 6** (Overflow): File larger than 2GB (Long limit exceeded)
//!
//! ## Performance Considerations
//!
//! - `FileLen` is very fast (reads file metadata only)
//! - Does not open the file or read contents
//! - Much faster than opening file to determine size
//! - Performance depends on file system and disk speed
//! - Network paths are slower than local paths
//! - Consider caching results if checking same file repeatedly
//!
//! ## Best Practices
//!
//! ### Check File Existence First
//!
//! ```vb
//! ' Good - Check existence to avoid error
//! If Dir(filePath) <> "" Then
//!     fileSize = FileLen(filePath)
//! Else
//!     MsgBox "File not found"
//! End If
//!
//! ' Or use error handling
//! On Error Resume Next
//! fileSize = FileLen(filePath)
//! If Err.Number <> 0 Then
//!     MsgBox "Cannot get file size"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Format for Display
//!
//! ```vb
//! ' Good - Format sizes for readability
//! Dim size As Long
//! size = FileLen(filePath)
//! MsgBox "File size: " & FormatFileSize(size)
//!
//! ' Bad - Raw bytes for large files
//! MsgBox "File size: " & FileLen(filePath) & " bytes"
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FileLen` vs `LOF`
//!
//! ```vb
//! ' FileLen - For closed files, returns current size
//! fileSize = FileLen("C:\data.txt")
//!
//! ' LOF - For open files only, returns current size
//! Open "C:\data.txt" For Input As #1
//! fileSize = LOF(1)
//! Close #1
//! ```
//!
//! ### `FileLen` vs `FileSystemObject.GetFile.Size`
//!
//! ```vb
//! ' FileLen - Built-in VB6 function
//! fileSize = FileLen("C:\data.txt")
//!
//! ' FSO - Requires reference to Scripting Runtime
//! Dim fso As Object
//! Set fso = CreateObject("Scripting.FileSystemObject")
//! fileSize = fso.GetFile("C:\data.txt").Size
//! ```
//!
//! ## Limitations
//!
//! - Maximum file size: 2,147,483,647 bytes (2GB - 1) due to `Long` type
//! - For files > 2GB, use `FileSystemObject` or API calls
//! - File must exist (cannot get size of non-existent files)
//! - Cannot get size of directories
//! - Returns size before opening for open files
//! - No built-in wildcard support (must use with `Dir`)
//!
//! ## Related Functions
//!
//! - `LOF`: Returns length of open file
//! - `Dir`: Returns file names matching a pattern
//! - `FileDateTime`: Returns file modification date/time
//! - `GetAttr`: Returns file attributes
//! - `FreeFile`: Returns next available file number
//! - `Open`: Opens a file for reading or writing

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn filelen_basic() {
        let source = r#"
fileSize = FileLen("C:\data.txt")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_variable() {
        let source = r#"
size = FileLen(filePath)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_debug_print() {
        let source = r#"
Debug.Print FileLen("C:\temp.dat")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_format() {
        let source = r#"
formatted = Format(FileLen(filePath) / 1024, "0.00") & " KB"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_in_function() {
        let source = r#"
Function GetFileSize(path As String) As Long
    GetFileSize = FileLen(path)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_addition() {
        let source = r#"
totalSize = totalSize + FileLen(fullPath)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_comparison() {
        let source = r#"
isLarger = (FileLen(file1) > FileLen(file2))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_if_statement() {
        let source = r#"
If FileLen(fullPath) > maxSize Then
    Debug.Print "File too large"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_error_handling() {
        let source = r#"
On Error Resume Next
size = FileLen(filePath)
If Err.Number <> 0 Then
    MsgBox "Error"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_loop() {
        let source = r#"
Do While fileName <> ""
    fileSize = FileLen(folderPath & fileName)
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_concatenation() {
        let source = r#"
msg = "Size: " & FileLen(filePath) & " bytes"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_range_check() {
        let source = r#"
If FileLen(fullPath) >= minSize And FileLen(fullPath) <= maxSize Then
    files.Add fullPath
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_equality() {
        let source = r#"
isValid = (FileLen(filePath) = expectedSize)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_zero_check() {
        let source = r#"
If FileLen(fullPath) = 0 Then
    emptyFiles.Add fullPath
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_udt_field() {
        let source = r#"
stats.TotalSize = stats.TotalSize + FileLen(fullPath)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_max_comparison() {
        let source = r#"
If FileLen(fullPath) > largestSize Then
    largestSize = FileLen(fullPath)
    largestFile = fileName
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_buffer_allocation() {
        let source = r#"
fileSize = FileLen(filePath)
ReDim buffer(0 To fileSize - 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_progress_calculation() {
        let source = r#"
ProgressBar.Value = (bytesRead / FileLen(sourceFile)) * 100
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_array_assignment() {
        let source = r#"
files(count).Size = FileLen(fullPath)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_division() {
        let source = r#"
averageSize = totalSize / fileCount
sizeInMB = FileLen(filePath) / 1048576
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_msgbox() {
        let source = r#"
MsgBox "File size: " & FormatFileSize(FileLen(filePath))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_select_case() {
        let source = r#"
Select Case FileLen(filePath)
    Case Is < 1024
        Debug.Print "Small"
    Case Is < 1048576
        Debug.Print "Medium"
    Case Else
        Debug.Print "Large"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_print_statement() {
        let source = r#"
Print #reportNum, fileName, FileLen(fullPath)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_for_loop() {
        let source = r#"
For i = 0 To fileCount - 1
    totalSize = totalSize + FileLen(files(i))
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_multiline() {
        let source = r#"
info = "File: " & fileName & vbCrLf & _
       "Size: " & FileLen(fullPath) & " bytes" & vbCrLf & _
       "Size (MB): " & Format(FileLen(fullPath) / 1048576, "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn filelen_category_check() {
        let source = r#"
If FileLen(filePath) < 102400 Then
    category = "Small"
ElseIf FileLen(filePath) < 10485760 Then
    category = "Medium"
Else
    category = "Large"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileLen"));
        assert!(debug.contains("Identifier"));
    }
}
