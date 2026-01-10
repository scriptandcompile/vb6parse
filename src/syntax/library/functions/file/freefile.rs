//! `FreeFile` Function
//!
//! Returns an `Integer` representing the next file number available for use by the Open statement.
//!
//! # Syntax
//!
//! ```vb
//! FreeFile[(rangenumber)]
//! ```
//!
//! # Parameters
//!
//! - `rangenumber` - Optional. Variant that specifies which range of file numbers to use.
//!   - `0` (default) - Returns a file number in the range 1-255 (inclusive).
//!   - `1` - Returns a file number in the range 256-511 (inclusive).
//!
//! # Return Value
//!
//! Returns an `Integer` representing the next available file number that is not already in use.
//!
//! # Remarks
//!
//! - Use `FreeFile` to obtain a file number that is not already associated with an open file.
//! - `FreeFile` returns the lowest available file number in the specified range.
//! - When using multiple files, always use `FreeFile` to avoid conflicts with file numbers.
//! - The file number returned can be used with the `Open` statement to open a file.
//! - After obtaining a file number with `FreeFile`, use it immediately to avoid conflicts.
//! - File numbers are released when the file is closed with the `Close` statement.
//! - The function is particularly important in libraries and reusable code where you don't know what file numbers are already in use.
//!
//! # Typical Uses
//!
//! - Opening files for sequential, random, or binary access
//! - Writing to log files
//! - Reading configuration files
//! - Temporary file operations
//! - Data import/export operations
//! - File-based data storage
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Basic file operations
//! Dim fileNum As Integer
//! fileNum = FreeFile
//! Open "data.txt" For Output As #fileNum
//! Print #fileNum, "Hello, World!"
//! Close #fileNum
//!
//! ' Multiple file operations
//! Dim inputFile As Integer
//! Dim outputFile As Integer
//! inputFile = FreeFile
//! Open "input.txt" For Input As #inputFile
//! outputFile = FreeFile
//! Open "output.txt" For Output As #outputFile
//! ' Process files...
//! Close #inputFile
//! Close #outputFile
//!
//! ' Using range parameter
//! Dim highRangeFile As Integer
//! highRangeFile = FreeFile(1)  ' Returns 256-511
//! Open "temp.dat" For Binary As #highRangeFile
//! Close #highRangeFile
//!
//! ' Immediate use pattern
//! Open "log.txt" For Append As #FreeFile
//! ' Note: Cannot close without saving the file number
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Simple File Read
//!
//! ```vb
//! Sub ReadTextFile(filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         Debug.Print line
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## 2. Simple File Write
//!
//! ```vb
//! Sub WriteTextFile(filename As String, content As String)
//!     Dim fileNum As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     Print #fileNum, content
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## 3. Append to Log File
//!
//! ```vb
//! Sub LogMessage(message As String)
//!     Dim fileNum As Integer
//!     
//!     fileNum = FreeFile
//!     Open App.Path & "\app.log" For Append As #fileNum
//!     Print #fileNum, Now & " - " & message
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## 4. Copy File Contents
//!
//! ```vb
//! Sub CopyFile(source As String, destination As String)
//!     Dim sourceNum As Integer
//!     Dim destNum As Integer
//!     Dim line As String
//!     
//!     sourceNum = FreeFile
//!     Open source For Input As #sourceNum
//!     
//!     destNum = FreeFile
//!     Open destination For Output As #destNum
//!     
//!     Do While Not EOF(sourceNum)
//!         Line Input #sourceNum, line
//!         Print #destNum, line
//!     Loop
//!     
//!     Close #sourceNum
//!     Close #destNum
//! End Sub
//! ```
//!
//! ## 5. Read Binary File
//!
//! ```vb
//! Function ReadBinaryFile(filename As String) As Byte()
//!     Dim fileNum As Integer
//!     Dim fileSize As Long
//!     Dim buffer() As Byte
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     
//!     fileSize = LOF(fileNum)
//!     ReDim buffer(0 To fileSize - 1)
//!     Get #fileNum, , buffer
//!     
//!     Close #fileNum
//!     ReadBinaryFile = buffer
//! End Function
//! ```
//!
//! ## 6. Write Binary File
//!
//! ```vb
//! Sub WriteBinaryFile(filename As String, data() As Byte)
//!     Dim fileNum As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     Put #fileNum, , data
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## 7. Read INI-Style Configuration
//!
//! ```vb
//! Function ReadConfigValue(filename As String, key As String) As String
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim pos As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         line = Trim(line)
//!         
//!         If Left(line, Len(key) + 1) = key & "=" Then
//!             ReadConfigValue = Mid(line, Len(key) + 2)
//!             Close #fileNum
//!             Exit Function
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     ReadConfigValue = ""
//! End Function
//! ```
//!
//! ## 8. Write CSV File
//!
//! ```vb
//! Sub WriteCSV(filename As String, data() As Variant)
//!     Dim fileNum As Integer
//!     Dim i As Long, j As Long
//!     Dim row As String
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     For i = LBound(data, 1) To UBound(data, 1)
//!         row = ""
//!         For j = LBound(data, 2) To UBound(data, 2)
//!             If j > LBound(data, 2) Then row = row & ","
//!             row = row & CStr(data(i, j))
//!         Next j
//!         Print #fileNum, row
//!     Next i
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## 9. Read Entire File into String
//!
//! ```vb
//! Function ReadFileAsString(filename As String) As String
//!     Dim fileNum As Integer
//!     Dim fileContent As String
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     fileContent = Input(LOF(fileNum), fileNum)
//!     Close #fileNum
//!     
//!     ReadFileAsString = fileContent
//! End Function
//! ```
//!
//! ## 10. Process Large File Line by Line
//!
//! ```vb
//! Sub ProcessLargeFile(filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim lineCount As Long
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     lineCount = 0
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         lineCount = lineCount + 1
//!         
//!         ' Process line here
//!         If lineCount Mod 1000 = 0 Then
//!             Debug.Print "Processed " & lineCount & " lines..."
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     Debug.Print "Total lines: " & lineCount
//! End Sub
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. File Manager Class
//!
//! ```vb
//! ' Class module: FileManager
//! Private m_FileNumber As Integer
//!
//! Public Sub OpenFile(filename As String, mode As String)
//!     m_FileNumber = FreeFile
//!     
//!     Select Case LCase(mode)
//!         Case "read"
//!             Open filename For Input As #m_FileNumber
//!         Case "write"
//!             Open filename For Output As #m_FileNumber
//!         Case "append"
//!             Open filename For Append As #m_FileNumber
//!         Case "binary"
//!             Open filename For Binary As #m_FileNumber
//!     End Select
//! End Sub
//!
//! Public Sub WriteLine(text As String)
//!     Print #m_FileNumber, text
//! End Sub
//!
//! Public Function ReadLine() As String
//!     Dim line As String
//!     If Not EOF(m_FileNumber) Then
//!         Line Input #m_FileNumber, line
//!         ReadLine = line
//!     End If
//! End Function
//!
//! Public Function IsEOF() As Boolean
//!     IsEOF = EOF(m_FileNumber)
//! End Function
//!
//! Public Sub CloseFile()
//!     If m_FileNumber > 0 Then
//!         Close #m_FileNumber
//!         m_FileNumber = 0
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseFile
//! End Sub
//! ```
//!
//! ## 2. Safe File Operations with Error Handling
//!
//! ```vb
//! Function SafeWriteFile(filename As String, content As String) As Boolean
//!     Dim fileNum As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     Print #fileNum, content
//!     Close #fileNum
//!     
//!     SafeWriteFile = True
//!     Exit Function
//!     
//! ErrorHandler:
//!     If fileNum > 0 Then Close #fileNum
//!     SafeWriteFile = False
//!     Debug.Print "Error writing file: " & Err.Description
//! End Function
//! ```
//!
//! ## 3. Multi-File Transaction Manager
//!
//! ```vb
//! Type FileHandle
//!     Number As Integer
//!     Filename As String
//!     IsOpen As Boolean
//! End Type
//!
//! Sub TransactionExample()
//!     Dim files(1 To 3) As FileHandle
//!     Dim i As Integer
//!     
//!     On Error GoTo Rollback
//!     
//!     ' Open multiple files
//!     For i = 1 To 3
//!         files(i).Number = FreeFile
//!         files(i).Filename = "file" & i & ".txt"
//!         files(i).IsOpen = False
//!     Next i
//!     
//!     ' Perform operations
//!     For i = 1 To 3
//!         Open files(i).Filename For Output As #files(i).Number
//!         files(i).IsOpen = True
//!         Print #files(i).Number, "Data for file " & i
//!     Next i
//!     
//!     ' Success - close all
//!     For i = 1 To 3
//!         Close #files(i).Number
//!         files(i).IsOpen = False
//!     Next i
//!     
//!     Exit Sub
//!     
//! Rollback:
//!     ' Clean up any open files
//!     For i = 1 To 3
//!         If files(i).IsOpen Then
//!             Close #files(i).Number
//!         End If
//!     Next i
//!     
//!     MsgBox "Transaction failed: " & Err.Description
//! End Sub
//! ```
//!
//! ## 4. File Pool Manager
//!
//! ```vb
//! ' Module-level variables
//! Private m_FilePool() As Integer
//! Private m_PoolSize As Integer
//!
//! Sub InitializeFilePool(poolSize As Integer)
//!     Dim i As Integer
//!     
//!     m_PoolSize = poolSize
//!     ReDim m_FilePool(1 To poolSize)
//!     
//!     For i = 1 To poolSize
//!         m_FilePool(i) = FreeFile(1)  ' Use high range
//!     Next i
//! End Sub
//!
//! Function GetPooledFileNumber() As Integer
//!     Static currentIndex As Integer
//!     
//!     currentIndex = currentIndex + 1
//!     If currentIndex > m_PoolSize Then currentIndex = 1
//!     
//!     GetPooledFileNumber = m_FilePool(currentIndex)
//! End Function
//! ```
//!
//! ## 5. Buffered File Writer
//!
//! ```vb
//! Type BufferedWriter
//!     FileNumber As Integer
//!     Buffer As String
//!     BufferSize As Long
//!     MaxBufferSize As Long
//! End Type
//!
//! Sub InitBufferedWriter(writer As BufferedWriter, filename As String, _
//!                        Optional maxBuffer As Long = 4096)
//!     writer.FileNumber = FreeFile
//!     writer.MaxBufferSize = maxBuffer
//!     writer.BufferSize = 0
//!     writer.Buffer = ""
//!     
//!     Open filename For Output As #writer.FileNumber
//! End Sub
//!
//! Sub BufferedWrite(writer As BufferedWriter, text As String)
//!     writer.Buffer = writer.Buffer & text & vbCrLf
//!     writer.BufferSize = Len(writer.Buffer)
//!     
//!     If writer.BufferSize >= writer.MaxBufferSize Then
//!         Print #writer.FileNumber, writer.Buffer;
//!         writer.Buffer = ""
//!         writer.BufferSize = 0
//!     End If
//! End Sub
//!
//! Sub FlushBufferedWriter(writer As BufferedWriter)
//!     If writer.BufferSize > 0 Then
//!         Print #writer.FileNumber, writer.Buffer;
//!         writer.Buffer = ""
//!         writer.BufferSize = 0
//!     End If
//!     
//!     Close #writer.FileNumber
//! End Sub
//! ```
//!
//! ## 6. Temporary File Manager
//!
//! ```vb
//! Function CreateTempFile() As Integer
//!     Dim fileNum As Integer
//!     Dim tempPath As String
//!     Dim tempFile As String
//!     
//!     tempPath = Environ("TEMP")
//!     If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
//!     
//!     tempFile = tempPath & "temp_" & Format(Now, "yyyymmddhhnnss") & ".tmp"
//!     
//!     fileNum = FreeFile
//!     Open tempFile For Binary As #fileNum
//!     
//!     CreateTempFile = fileNum
//! End Function
//!
//! Sub DeleteTempFile(fileNum As Integer)
//!     Dim filename As String
//!     
//!     ' Get filename before closing
//!     filename = FileAttr(fileNum, 1)  ' Returns filename
//!     
//!     Close #fileNum
//!     Kill filename  ' Delete the file
//! End Sub
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeFileOperation(filename As String) As Boolean
//!     Dim fileNum As Integer
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     ' Process file...
//!     
//!     Close #fileNum
//!     SafeFileOperation = True
//!     Exit Function
//!     
//! ErrorHandler:
//!     ' Always try to close the file
//!     On Error Resume Next
//!     If fileNum > 0 Then Close #fileNum
//!     On Error GoTo 0
//!     
//!     Select Case Err.Number
//!         Case 53  ' File not found
//!             Debug.Print "File not found: " & filename
//!         Case 55  ' File already open
//!             Debug.Print "File already open: " & filename
//!         Case 70  ' Permission denied
//!             Debug.Print "Permission denied: " & filename
//!         Case 71  ' Disk not ready
//!             Debug.Print "Disk not ready"
//!         Case 76  ' Path not found
//!             Debug.Print "Path not found: " & filename
//!         Case Else
//!             Debug.Print "Error " & Err.Number & ": " & Err.Description
//!     End Select
//!     
//!     SafeFileOperation = False
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 55 (File already open)**: File number is already in use. Always use `FreeFile` to avoid this.
//! - **Error 52 (Bad file name or number)**: Invalid file number. Ensure the number is in the valid range.
//! - **Error 53 (File not found)**: The specified file does not exist.
//! - **Error 76 (Path not found)**: The specified path does not exist.
//!
//! # Performance Considerations
//!
//! - `FreeFile` is a very fast operation - no overhead in calling it
//! - Always store the result in a variable for later use with `Close`
//! - Don't call `FreeFile` repeatedly in tight loops - get the number once and reuse it
//! - Consider using the high range (256-511) for system or library files to avoid conflicts with user code
//! - File operations themselves (`Open`, `Close`, `Read`, `Write`) are much slower than `FreeFile`
//!
//! # Best Practices
//!
//! 1. **Always use `FreeFile`** instead of hard-coding file numbers
//! 2. **Store the file number** in a variable so you can close the file later
//! 3. **Close files promptly** when done to release the file number
//! 4. **Use error handling** to ensure files are closed even if an error occurs
//! 5. **Use high range (1)** for system/library code to avoid conflicts
//! 6. **Open and close files in pairs** - every `Open` should have a corresponding `Close`
//! 7. **Don't assume file numbers** - always get a fresh number with `FreeFile`
//!
//! # Comparison with Other Approaches
//!
//! ## `FreeFile` vs Hard-Coded Numbers
//!
//! ```vb
//! ' Bad - Hard-coded file number
//! Open "data.txt" For Input As #1
//! ' What if file #1 is already open?
//!
//! ' Good - Use FreeFile
//! Dim fileNum As Integer
//! fileNum = FreeFile
//! Open "data.txt" For Input As #fileNum
//! ```
//!
//! ## `FreeFile` vs `FileSystemObject`
//!
//! ```vb
//! ' FreeFile approach - built-in, fast, simple
//! Dim fileNum As Integer
//! fileNum = FreeFile
//! Open "data.txt" For Input As #fileNum
//! ' ...
//! Close #fileNum
//!
//! ' FileSystemObject - more features, requires reference
//! Dim fso As New FileSystemObject
//! Dim ts As TextStream
//! Set ts = fso.OpenTextFile("data.txt", ForReading)
//! ' ...
//! ts.Close
//! Set ts = Nothing
//! Set fso = Nothing
//! ```
//!
//! # Limitations
//!
//! - Maximum of 255 files in the default range (1-255)
//! - Maximum of 256 files in the high range (256-511)
//! - Total maximum of 511 files open simultaneously
//! - File numbers are process-specific (not thread-safe in modern contexts)
//! - No built-in file locking mechanism
//! - Traditional file I/O is slower than modern streaming APIs
//!
//! # Range Selection
//!
//! When to use different ranges:
//!
//! - **Range 0 (1-255)**: Default for application code, user files
//! - **Range 1 (256-511)**: System files, library code, background operations
//!
//! ```vb
//! ' Application uses default range
//! Dim userFile As Integer
//! userFile = FreeFile(0)  ' or just FreeFile
//!
//! ' Library uses high range to avoid conflicts
//! Dim sysFile As Integer
//! sysFile = FreeFile(1)
//! ```
//!
//! # Related Functions
//!
//! - `Open` - Opens a file for reading or writing
//! - `Close` - Closes an open file
//! - `EOF` - Tests for end-of-file
//! - `LOF` - Returns the length of an open file
//! - `FileAttr` - Returns file mode or file handle
//! - `Seek` - Sets or returns the current file position
//! - `Input` - Reads data from a sequential file
//! - `Print` - Writes data to a sequential file
//! - `Get` - Reads data from a binary or random file
//! - `Put` - Writes data to a binary or random file

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn freefile_basic() {
        let source = r"fileNum = FreeFile";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_parentheses() {
        let source = r"fileNum = FreeFile()";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_range() {
        let source = r"fileNum = FreeFile(1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_range_zero() {
        let source = r"fileNum = FreeFile(0)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_in_open() {
        let source = r#"Open "data.txt" For Output As #FreeFile"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_assignment() {
        let source = r"Dim fileNum As Integer
fileNum = FreeFile";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_multiple() {
        let source = r"inputFile = FreeFile
outputFile = FreeFile";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_in_function() {
        let source = r"Function ReadFile() As String
    Dim fileNum As Integer
    fileNum = FreeFile
End Function";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_in_sub() {
        let source = r#"Sub WriteLog()
    Dim fileNum As Integer
    fileNum = FreeFile
    Open "log.txt" For Append As #fileNum
End Sub"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_open_input() {
        let source = r"fileNum = FreeFile
Open filename For Input As #fileNum";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_open_binary() {
        let source = r"fileNum = FreeFile
Open filename For Binary As #fileNum";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_error_handling() {
        let source = r"On Error GoTo ErrorHandler
fileNum = FreeFile
Open filename For Input As #fileNum";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_in_loop() {
        let source = r"For i = 1 To 3
    files(i) = FreeFile
Next i";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_array_assignment() {
        let source = r"m_FilePool(i) = FreeFile(1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_type_member() {
        let source = r"writer.FileNumber = FreeFile";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_if_statement() {
        let source = r"If fileNum = 0 Then fileNum = FreeFile";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_comparison() {
        let source = r#"If FreeFile > 255 Then MsgBox "High range""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_debug_print() {
        let source = r#"Debug.Print "File number: " & FreeFile"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_class_initialize() {
        let source = r"Private Sub Class_Initialize()
    m_FileNumber = FreeFile
End Sub";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_close() {
        let source = r#"fileNum = FreeFile
Open "data.txt" For Input As #fileNum
Close #fileNum"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_do_loop() {
        let source = r"fileNum = FreeFile
Open filename For Input As #fileNum
Do While Not EOF(fileNum)
    Line Input #fileNum, line
Loop
Close #fileNum";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_select_case() {
        let source = r#"Select Case mode
    Case "read"
        fileNum = FreeFile
        Open filename For Input As #fileNum
End Select"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_function_return() {
        let source = r"Function GetFileNum() As Integer
    GetFileNum = FreeFile
End Function";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_concatenation() {
        let source = r#"msg = "File: " & filename & " Num: " & FreeFile"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_with_lof() {
        let source = r"fileNum = FreeFile
Open filename For Binary As #fileNum
fileSize = LOF(fileNum)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn freefile_temp_file() {
        let source = r"tempFile = CreateTempFile()
Function CreateTempFile() As Integer
    CreateTempFile = FreeFile
End Function";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/file/freefile",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
