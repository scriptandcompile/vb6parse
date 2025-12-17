//! # `FileAttr` Function
//!
//! Returns a `Long` representing the file mode for files opened using the `Open` statement,
//! or the file attribute information for files, directories, or folders.
//!
//! ## Syntax
//!
//! ```vb
//! FileAttr(filenumber, returntype)
//! ```
//!
//! ## Parameters
//!
//! - **filenumber**: Required. An `Integer` containing a valid file number of an open file.
//! - **returntype**: Required. A `Long` indicating the type of information to return.
//!   - **1**: Returns a value indicating the file mode (`Input`, `Output`, `Append`, `Binary`, `Random`)
//!   - **2**: Returns the file handle used by the operating system
//!
//! ## Return Value
//!
//! Returns a `Long` value. The meaning depends on the returntype parameter:
//!
//! ### When returntype = 1 (File Mode):
//! - **1**: `Input` mode
//! - **2**: `Output` mode
//! - **4**: `Random` access mode
//! - **8**: `Append` mode
//! - **32**: `Binary` mode
//!
//! ### When returntype = 2 (File Handle):
//! Returns the operating system file handle (an integer value used by the OS to identify the file).
//!
//! ## Remarks
//!
//! The `FileAttr` function returns information about files that have been opened using
//! the `Open` statement. It provides two types of information: the file access mode
//! or the operating system file handle.
//!
//! **Important Characteristics:**
//!
//! - File must be open before calling `FileAttr`
//! - Error if file number is invalid or file is closed
//! - returntype must be 1 or 2
//! - File mode values are mutually exclusive
//! - File handle is OS-specific (Windows, Unix, etc.)
//! - File handle can be used with API calls
//! - Not applicable to files opened by other applications
//! - Only works with files opened via VB6 Open statement
//!
//! ## File Mode Values (returntype = 1)
//!
//! | Mode | Value | Description |
//! |------|-------|-------------|
//! | Input | 1 | File opened for reading |
//! | Output | 2 | File opened for writing (new file or overwrite) |
//! | Random | 4 | File opened for random access |
//! | Append | 8 | File opened for appending |
//! | Binary | 32 | File opened in binary mode |
//!
//! ## Examples
//!
//! ### Basic Usage - Check File Mode
//!
//! ```vb
//! Sub CheckFileMode()
//!     Dim fileNum As Integer
//!     Dim fileMode As Long
//!     
//!     fileNum = FreeFile
//!     Open "C:\data.txt" For Input As #fileNum
//!     
//!     ' Get file mode
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     Select Case fileMode
//!         Case 1
//!             Debug.Print "File opened for Input"
//!         Case 2
//!             Debug.Print "File opened for Output"
//!         Case 4
//!             Debug.Print "File opened for Random access"
//!         Case 8
//!             Debug.Print "File opened for Append"
//!         Case 32
//!             Debug.Print "File opened for Binary"
//!     End Select
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Get File Handle
//!
//! ```vb
//! Sub GetFileHandle()
//!     Dim fileNum As Integer
//!     Dim fileHandle As Long
//!     
//!     fileNum = FreeFile
//!     Open "C:\temp.dat" For Binary As #fileNum
//!     
//!     ' Get operating system file handle
//!     fileHandle = FileAttr(fileNum, 2)
//!     Debug.Print "OS File Handle: " & fileHandle
//!     
//!     ' File handle can be used with Windows API calls
//!     ' Example: SetFilePointer, ReadFile, WriteFile, etc.
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Verify File Is Open for Writing
//!
//! ```vb
//! Function CanWriteToFile(fileNum As Integer) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileMode As Long
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     ' Check if file is open for Output, Append, Random, or Binary
//!     CanWriteToFile = (fileMode = 2 Or fileMode = 8 Or fileMode = 4 Or fileMode = 32)
//!     Exit Function
//!     
//! ErrorHandler:
//!     CanWriteToFile = False
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### File Mode Lookup Function
//!
//! ```vb
//! Function GetFileModeDescription(fileNum As Integer) As String
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileMode As Long
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     Select Case fileMode
//!         Case 1
//!             GetFileModeDescription = "Input"
//!         Case 2
//!             GetFileModeDescription = "Output"
//!         Case 4
//!             GetFileModeDescription = "Random"
//!         Case 8
//!             GetFileModeDescription = "Append"
//!         Case 32
//!             GetFileModeDescription = "Binary"
//!         Case Else
//!             GetFileModeDescription = "Unknown"
//!     End Select
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileModeDescription = "Error: " & Err.Description
//! End Function
//! ```
//!
//! ### List All Open Files
//!
//! ```vb
//! Sub ListOpenFiles()
//!     Dim i As Integer
//!     Dim fileMode As Long
//!     Dim modeDesc As String
//!     
//!     Debug.Print "Open Files:"
//!     Debug.Print String(60, "-")
//!     Debug.Print "File#", "Mode", "Handle"
//!     Debug.Print String(60, "-")
//!     
//!     For i = 1 To 255
//!         On Error Resume Next
//!         fileMode = FileAttr(i, 1)
//!         
//!         If Err.Number = 0 Then
//!             ' File is open
//!             modeDesc = GetFileModeDescription(i)
//!             Debug.Print i, modeDesc, FileAttr(i, 2)
//!         End If
//!         
//!         Err.Clear
//!     Next i
//! End Sub
//! ```
//!
//! ### Safe File Operation Wrapper
//!
//! ```vb
//! Function WriteToFile(fileNum As Integer, data As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileMode As Long
//!     
//!     ' Verify file is open
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     ' Check if writable
//!     If fileMode <> 2 And fileMode <> 8 And fileMode <> 4 And fileMode <> 32 Then
//!         MsgBox "File is not open for writing", vbExclamation
//!         WriteToFile = False
//!         Exit Function
//!     End If
//!     
//!     ' Write data based on mode
//!     Select Case fileMode
//!         Case 2, 8  ' Output or Append
//!             Print #fileNum, data
//!         Case 4     ' Random
//!             Put #fileNum, , data
//!         Case 32    ' Binary
//!             Put #fileNum, , data
//!     End Select
//!     
//!     WriteToFile = True
//!     Exit Function
//!     
//! ErrorHandler:
//!     WriteToFile = False
//! End Function
//! ```
//!
//! ### Log File Access Information
//!
//! ```vb
//! Sub LogFileAccess(fileNum As Integer, logPath As String)
//!     Dim logNum As Integer
//!     Dim fileMode As Long
//!     Dim fileHandle As Long
//!     
//!     On Error Resume Next
//!     
//!     fileMode = FileAttr(fileNum, 1)
//!     fileHandle = FileAttr(fileNum, 2)
//!     
//!     If Err.Number = 0 Then
//!         logNum = FreeFile
//!         Open logPath For Append As #logNum
//!         Print #logNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
//!                        "File#" & fileNum & " | " & _
//!                        "Mode: " & GetFileModeDescription(fileNum) & " | " & _
//!                        "Handle: " & fileHandle
//!         Close #logNum
//!     End If
//! End Sub
//! ```
//!
//! ### Check If File Number Is Valid
//!
//! ```vb
//! Function IsFileOpen(fileNum As Integer) As Boolean
//!     On Error Resume Next
//!     Dim fileMode As Long
//!     
//!     fileMode = FileAttr(fileNum, 1)
//!     IsFileOpen = (Err.Number = 0)
//! End Function
//! ```
//!
//! ### Get All File Handles
//!
//! ```vb
//! Function GetOpenFileHandles() As Collection
//!     Dim handles As New Collection
//!     Dim i As Integer
//!     Dim fileHandle As Long
//!     
//!     For i = 1 To 255
//!         On Error Resume Next
//!         fileHandle = FileAttr(i, 2)
//!         
//!         If Err.Number = 0 Then
//!             handles.Add fileHandle, CStr(i)
//!         End If
//!         
//!         Err.Clear
//!     Next i
//!     
//!     Set GetOpenFileHandles = handles
//! End Function
//! ```
//!
//! ### Validate File Before Operation
//!
//! ```vb
//! Function ValidateFileForReading(fileNum As Integer) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileMode As Long
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     ' Check if file is open for Input, Random, or Binary
//!     ValidateFileForReading = (fileMode = 1 Or fileMode = 4 Or fileMode = 32)
//!     Exit Function
//!     
//! ErrorHandler:
//!     ValidateFileForReading = False
//! End Function
//! ```
//!
//! ### Compare File Modes
//!
//! ```vb
//! Sub CompareFileModes(file1 As Integer, file2 As Integer)
//!     Dim mode1 As Long, mode2 As Long
//!     
//!     On Error Resume Next
//!     mode1 = FileAttr(file1, 1)
//!     mode2 = FileAttr(file2, 1)
//!     
//!     If Err.Number = 0 Then
//!         If mode1 = mode2 Then
//!             Debug.Print "Files have the same mode: " & GetFileModeDescription(file1)
//!         Else
//!             Debug.Print "File1 mode: " & GetFileModeDescription(file1)
//!             Debug.Print "File2 mode: " & GetFileModeDescription(file2)
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ### Track File Usage Statistics
//!
//! ```vb
//! Type FileStats
//!     FileNumber As Integer
//!     Mode As Long
//!     Handle As Long
//!     OpenTime As Date
//!     OperationCount As Long
//! End Type
//!
//! Private fileStatistics() As FileStats
//! Private statCount As Long
//!
//! Sub RecordFileOpen(fileNum As Integer)
//!     On Error Resume Next
//!     
//!     Dim fileMode As Long
//!     Dim fileHandle As Long
//!     
//!     fileMode = FileAttr(fileNum, 1)
//!     fileHandle = FileAttr(fileNum, 2)
//!     
//!     If Err.Number = 0 Then
//!         ReDim Preserve fileStatistics(0 To statCount)
//!         
//!         With fileStatistics(statCount)
//!             .FileNumber = fileNum
//!             .Mode = fileMode
//!             .Handle = fileHandle
//!             .OpenTime = Now
//!             .OperationCount = 0
//!         End With
//!         
//!         statCount = statCount + 1
//!     End If
//! End Sub
//! ```
//!
//! ### Platform-Specific File Handle Usage
//!
//! ```vb
//! ' Windows API declarations (for demonstration)
//! Private Declare Function GetFileSize Lib "kernel32" _
//!     (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
//!
//! Function GetFileSizeViaHandle(fileNum As Integer) As Long
//!     Dim fileHandle As Long
//!     Dim fileSizeHigh As Long
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Get file handle from VB6 file number
//!     fileHandle = FileAttr(fileNum, 2)
//!     
//!     ' Use Windows API to get file size
//!     GetFileSizeViaHandle = GetFileSize(fileHandle, fileSizeHigh)
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileSizeViaHandle = -1
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### File Access Monitor
//!
//! ```vb
//! Type FileAccessInfo
//!     FileNumber As Integer
//!     Mode As String
//!     Handle As Long
//!     CanRead As Boolean
//!     CanWrite As Boolean
//!     LastChecked As Date
//! End Type
//!
//! Function GetFileAccessInfo(fileNum As Integer) As FileAccessInfo
//!     Dim info As FileAccessInfo
//!     Dim fileMode As Long
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     With info
//!         .FileNumber = fileNum
//!         .Handle = FileAttr(fileNum, 2)
//!         .LastChecked = Now
//!         
//!         Select Case fileMode
//!             Case 1  ' Input
//!                 .Mode = "Input"
//!                 .CanRead = True
//!                 .CanWrite = False
//!             Case 2  ' Output
//!                 .Mode = "Output"
//!                 .CanRead = False
//!                 .CanWrite = True
//!             Case 4  ' Random
//!                 .Mode = "Random"
//!                 .CanRead = True
//!                 .CanWrite = True
//!             Case 8  ' Append
//!                 .Mode = "Append"
//!                 .CanRead = False
//!                 .CanWrite = True
//!             Case 32  ' Binary
//!                 .Mode = "Binary"
//!                 .CanRead = True
//!                 .CanWrite = True
//!         End Select
//!     End With
//!     
//!     GetFileAccessInfo = info
//!     Exit Function
//!     
//! ErrorHandler:
//!     info.Mode = "Error"
//!     GetFileAccessInfo = info
//! End Function
//! ```
//!
//! ### Automatic File Mode Detection for Operations
//!
//! ```vb
//! Function ReadFromFile(fileNum As Integer, ByRef data As Variant) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileMode As Long
//!     fileMode = FileAttr(fileNum, 1)
//!     
//!     ' Read based on detected mode
//!     Select Case fileMode
//!         Case 1  ' Input mode
//!             If Not EOF(fileNum) Then
//!                 Line Input #fileNum, data
//!                 ReadFromFile = True
//!             End If
//!         
//!         Case 4  ' Random mode
//!             Get #fileNum, , data
//!             ReadFromFile = True
//!         
//!         Case 32  ' Binary mode
//!             Get #fileNum, , data
//!             ReadFromFile = True
//!         
//!         Case Else
//!             MsgBox "File not open for reading", vbExclamation
//!             ReadFromFile = False
//!     End Select
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     ReadFromFile = False
//! End Function
//! ```
//!
//! ### File Handle Cache
//!
//! ```vb
//! Private Type HandleCacheEntry
//!     FileNumber As Integer
//!     Handle As Long
//!     CachedTime As Date
//! End Type
//!
//! Private handleCache() As HandleCacheEntry
//! Private cacheSize As Long
//!
//! Function GetCachedHandle(fileNum As Integer) As Long
//!     Dim i As Long
//!     Dim currentTime As Date
//!     
//!     currentTime = Now
//!     
//!     ' Check cache first
//!     For i = 0 To cacheSize - 1
//!         If handleCache(i).FileNumber = fileNum Then
//!             ' Verify cache is still valid (within 1 second)
//!             If DateDiff("s", handleCache(i).CachedTime, currentTime) < 1 Then
//!                 GetCachedHandle = handleCache(i).Handle
//!                 Exit Function
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Not in cache or expired, get fresh value
//!     On Error Resume Next
//!     GetCachedHandle = FileAttr(fileNum, 2)
//!     
//!     If Err.Number = 0 Then
//!         ' Add to cache
//!         ReDim Preserve handleCache(0 To cacheSize)
//!         handleCache(cacheSize).FileNumber = fileNum
//!         handleCache(cacheSize).Handle = GetCachedHandle
//!         handleCache(cacheSize).CachedTime = currentTime
//!         cacheSize = cacheSize + 1
//!     End If
//! End Function
//! ```
//!
//! ### Cross-Platform File Handle Wrapper
//!
//! ```vb
//! Function GetPlatformFileInfo(fileNum As Integer) As String
//!     Dim fileHandle As Long
//!     Dim info As String
//!     
//!     On Error Resume Next
//!     
//!     fileHandle = FileAttr(fileNum, 2)
//!     
//!     If Err.Number = 0 Then
//!         info = "File Number: " & fileNum & vbCrLf
//!         info = info & "Mode: " & GetFileModeDescription(fileNum) & vbCrLf
//!         info = info & "OS Handle: " & fileHandle & vbCrLf
//!         
//!         ' Platform-specific information
//!         #If Win32 Then
//!             info = info & "Platform: Windows 32-bit" & vbCrLf
//!         #ElseIf Win64 Then
//!             info = info & "Platform: Windows 64-bit" & vbCrLf
//!         #Else
//!             info = info & "Platform: Unknown" & vbCrLf
//!         #End If
//!     Else
//!         info = "File not open"
//!     End If
//!     
//!     GetPlatformFileInfo = info
//! End Function
//! ```
//!
//! ### File Descriptor Manager
//!
//! ```vb
//! Private Type FileDescriptor
//!     VBFileNumber As Integer
//!     OSHandle As Long
//!     Mode As Long
//!     ModeDescription As String
//!     FilePath As String
//!     OpenedAt As Date
//!     IsOpen As Boolean
//! End Type
//!
//! Private descriptors As Collection
//!
//! Sub InitializeDescriptorManager()
//!     Set descriptors = New Collection
//! End Sub
//!
//! Sub RegisterOpenFile(fileNum As Integer, filePath As String)
//!     Dim desc As FileDescriptor
//!     
//!     On Error Resume Next
//!     
//!     desc.VBFileNumber = fileNum
//!     desc.OSHandle = FileAttr(fileNum, 2)
//!     desc.Mode = FileAttr(fileNum, 1)
//!     desc.ModeDescription = GetFileModeDescription(fileNum)
//!     desc.FilePath = filePath
//!     desc.OpenedAt = Now
//!     desc.IsOpen = (Err.Number = 0)
//!     
//!     If desc.IsOpen Then
//!         descriptors.Add desc, "FD" & fileNum
//!     End If
//! End Sub
//!
//! Function GetDescriptor(fileNum As Integer) As FileDescriptor
//!     On Error Resume Next
//!     GetDescriptor = descriptors("FD" & fileNum)
//! End Function
//! ```
//!
//! ### Comprehensive File State Checker
//!
//! ```vb
//! Function GetCompleteFileState(fileNum As Integer) As String
//!     Dim state As String
//!     Dim fileMode As Long
//!     Dim fileHandle As Long
//!     
//!     On Error Resume Next
//!     
//!     state = "=== File State for #" & fileNum & " ===" & vbCrLf
//!     
//!     fileMode = FileAttr(fileNum, 1)
//!     If Err.Number <> 0 Then
//!         state = state & "File is CLOSED or invalid file number" & vbCrLf
//!         GetCompleteFileState = state
//!         Exit Function
//!     End If
//!     
//!     state = state & "File is OPEN" & vbCrLf
//!     state = state & "Mode: " & GetFileModeDescription(fileNum) & " (" & fileMode & ")" & vbCrLf
//!     
//!     fileHandle = FileAttr(fileNum, 2)
//!     state = state & "OS Handle: " & fileHandle & vbCrLf
//!     
//!     ' Add capabilities
//!     state = state & "Capabilities:" & vbCrLf
//!     Select Case fileMode
//!         Case 1
//!             state = state & "  - Read: Yes" & vbCrLf
//!             state = state & "  - Write: No" & vbCrLf
//!             state = state & "  - EOF applicable: Yes" & vbCrLf
//!         Case 2
//!             state = state & "  - Read: No" & vbCrLf
//!             state = state & "  - Write: Yes" & vbCrLf
//!             state = state & "  - EOF applicable: No" & vbCrLf
//!         Case 4
//!             state = state & "  - Read: Yes" & vbCrLf
//!             state = state & "  - Write: Yes" & vbCrLf
//!             state = state & "  - EOF applicable: Yes" & vbCrLf
//!         Case 8
//!             state = state & "  - Read: No" & vbCrLf
//!             state = state & "  - Write: Yes (append only)" & vbCrLf
//!             state = state & "  - EOF applicable: No" & vbCrLf
//!         Case 32
//!             state = state & "  - Read: Yes" & vbCrLf
//!             state = state & "  - Write: Yes" & vbCrLf
//!             state = state & "  - EOF applicable: Use LOF/Seek" & vbCrLf
//!     End Select
//!     
//!     GetCompleteFileState = state
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFileAttr(fileNum As Integer, returnType As Long) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate returnType
//!     If returnType <> 1 And returnType <> 2 Then
//!         Err.Raise 5, , "Invalid returnType. Must be 1 or 2."
//!     End If
//!     
//!     SafeFileAttr = FileAttr(fileNum, returnType)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 52  ' Bad file name or number
//!             MsgBox "File #" & fileNum & " is not open", vbExclamation
//!             SafeFileAttr = Null
//!         Case 5   ' Invalid procedure call
//!             MsgBox "Invalid returnType parameter", vbExclamation
//!             SafeFileAttr = Null
//!         Case Else
//!             MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
//!             SafeFileAttr = Null
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 52** (Bad file name or number): File number is invalid or file is closed
//! - **Error 5** (Invalid procedure call): returntype is not 1 or 2
//!
//! ## Performance Considerations
//!
//! - `FileAttr` is very fast (simple state query)
//! - Minimal overhead for checking file state
//! - More efficient than maintaining separate state variables
//! - File handle retrieval (returntype=2) is as fast as mode retrieval
//! - Consider caching results if calling frequently in tight loops
//!
//! ## Best Practices
//!
//! ### Always Validate File Is Open
//!
//! ```vb
//! ' Good - Check before operations
//! On Error Resume Next
//! fileMode = FileAttr(fileNum, 1)
//! If Err.Number <> 0 Then
//!     MsgBox "File is not open"
//!     Exit Sub
//! End If
//! On Error GoTo 0
//!
//! ' Or use IsFileOpen helper
//! If Not IsFileOpen(fileNum) Then
//!     Exit Sub
//! End If
//! ```
//!
//! ### Use Constants for Return Types
//!
//! ```vb
//! ' Good - Define constants for clarity
//! Const FILE_ATTR_MODE = 1
//! Const FILE_ATTR_HANDLE = 2
//!
//! fileMode = FileAttr(fileNum, FILE_ATTR_MODE)
//! fileHandle = FileAttr(fileNum, FILE_ATTR_HANDLE)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FileAttr` vs `GetAttr`
//!
//! ```vb
//! ' FileAttr - For open files, returns mode or handle
//! fileMode = FileAttr(fileNum, 1)  ' File must be open
//!
//! ' GetAttr - For any file, returns attributes (readonly, hidden, etc.)
//! attrs = GetAttr("C:\file.txt")   ' File can be closed
//! ```
//!
//! ### `FileAttr` vs `LOF`
//!
//! ```vb
//! ' FileAttr - Returns mode or handle
//! fileMode = FileAttr(fileNum, 1)
//!
//! ' LOF - Returns file length in bytes
//! fileSize = LOF(fileNum)
//! ```
//!
//! ## Limitations
//!
//! - Only works with files opened via VB6 Open statement
//! - Cannot get attributes of closed files
//! - returntype must be exactly 1 or 2
//! - File handle is platform-specific
//! - No information about file path or name
//! - Cannot determine if file is at EOF
//! - Does not indicate file position
//!
//! ## Related Functions
//!
//! - `FreeFile`: Returns next available file number
//! - `Open`: Opens a file for reading or writing
//! - `Close`: Closes an open file
//! - `LOF`: Returns length of open file
//! - `Seek`: Returns or sets current position in file
//! - `EOF`: Returns whether end of file reached
//! - `GetAttr`: Returns attributes of any file (readonly, hidden, etc.)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn fileattr_mode() {
        let source = r#"
fileMode = FileAttr(fileNum, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_handle() {
        let source = r#"
fileHandle = FileAttr(fileNum, 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_literal_file_number() {
        let source = r#"
mode = FileAttr(1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_select_case() {
        let source = r#"
Select Case FileAttr(fileNum, 1)
    Case 1
        Debug.Print "Input"
    Case 2
        Debug.Print "Output"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_in_if() {
        let source = r#"
If FileAttr(fileNum, 1) = 1 Then
    Debug.Print "Input mode"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_in_function() {
        let source = r#"
Function GetFileMode(fnum As Integer) As Long
    GetFileMode = FileAttr(fnum, 1)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_debug_print() {
        let source = r#"
Debug.Print FileAttr(fileNum, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_comparison() {
        let source = r#"
canWrite = (FileAttr(fileNum, 1) = 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_or_condition() {
        let source = r#"
isWritable = (FileAttr(fileNum, 1) = 2 Or FileAttr(fileNum, 1) = 8)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_error_handling() {
        let source = r#"
On Error Resume Next
mode = FileAttr(fileNum, 1)
If Err.Number <> 0 Then
    MsgBox "File not open"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_in_loop() {
        let source = r#"
For i = 1 To 255
    mode = FileAttr(i, 1)
    If Err.Number = 0 Then
        Debug.Print i, mode
    End If
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_with_concatenation() {
        let source = r#"
msg = "Mode: " & FileAttr(fileNum, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_udt_assignment() {
        let source = r#"
info.Mode = FileAttr(fileNum, 1)
info.Handle = FileAttr(fileNum, 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_validation() {
        let source = r#"
valid = (FileAttr(fileNum, 1) >= 1 And FileAttr(fileNum, 1) <= 32)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_nested_if() {
        let source = r#"
If FileAttr(fileNum, 1) = 1 Or FileAttr(fileNum, 1) = 4 Then
    canRead = True
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_collection_add() {
        let source = r#"
handles.Add FileAttr(i, 2), CStr(i)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_msgbox() {
        let source = r#"
MsgBox "File mode: " & FileAttr(fileNum, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_multiline() {
        let source = r#"
info = "File #" & fileNum & vbCrLf & _
       "Mode: " & FileAttr(fileNum, 1) & vbCrLf & _
       "Handle: " & FileAttr(fileNum, 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_format() {
        let source = r#"
formatted = Format(FileAttr(fileNum, 1), "0")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_file_logging() {
        let source = r#"
Print #logNum, "Mode: " & FileAttr(fileNum, 1) & " | Handle: " & FileAttr(fileNum, 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_boolean_expression() {
        let source = r#"
isOpen = (Err.Number = 0) And (FileAttr(fileNum, 1) > 0)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_with_constants() {
        let source = r#"
Const FILE_ATTR_MODE = 1
mode = FileAttr(fileNum, FILE_ATTR_MODE)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_array_assignment() {
        let source = r#"
modes(i) = FileAttr(i, 1)
handles(i) = FileAttr(i, 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_in_with_block() {
        let source = r#"
With descriptor
    .Mode = FileAttr(fileNum, 1)
    .Handle = FileAttr(fileNum, 2)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn fileattr_immediate_window() {
        let source = r#"
? FileAttr(1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FileAttr"));
        assert!(debug.contains("Identifier"));
    }
}
