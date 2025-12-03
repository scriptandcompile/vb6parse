//! # LOF Function
//!
//! Returns a Long representing the size, in bytes, of a file opened using the Open statement.
//!
//! ## Syntax
//!
//! ```vb
//! LOF(filenumber)
//! ```
//!
//! ## Parameters
//!
//! - `filenumber` (Required): Integer file number used in the Open statement
//!   - Must be a valid file number from a currently open file
//!   - File numbers typically obtained from `FreeFile` function
//!   - Must be between 1 and 511
//!
//! ## Return Value
//!
//! Returns a Long:
//! - Size of the file in bytes
//! - For files opened in any mode (Binary, Random, Input, Output, Append)
//! - Returns 0 for empty files
//! - Maximum value 2,147,483,647 (Long type limit ~2GB)
//! - Returns actual file size on disk
//! - Updated immediately if file grows during operation
//!
//! ## Remarks
//!
//! The LOF function returns the length (size) of an open file:
//!
//! - Works with all file access modes (Binary, Random, Input, Output, Append)
//! - Returns size in bytes regardless of mode
//! - File must be open before calling LOF
//! - Does not change file pointer position
//! - Read-only operation (non-destructive)
//! - Useful for calculating progress during file operations
//! - Essential for determining number of records in Random files
//! - Used to allocate buffers for reading entire file
//! - Can be used with Loc to calculate percentage complete
//! - Error 52 "Bad file name or number" if file not open
//! - Error 68 "Device unavailable" if device unavailable
//! - For Random files, divide by record length to get record count
//! - Returns current size, even if file is being written to
//! - More reliable than `FileLen` for open files
//! - `FileLen` works on closed files, LOF works on open files
//! - Common in loops reading files to completion
//! - Used to detect empty files (LOF returns 0)
//! - Essential for progress bars and status indicators
//! - Helps prevent reading past end of file
//!
//! ## Typical Uses
//!
//! 1. **Get File Size**
//!    ```vb
//!    fileSize = LOF(1)
//!    ```
//!
//! 2. **Calculate Record Count**
//!    ```vb
//!    recordCount = LOF(fileNum) / Len(record)
//!    ```
//!
//! 3. **Read Entire File**
//!    ```vb
//!    buffer = String(LOF(fileNum), 0)
//!    Get #fileNum, , buffer
//!    ```
//!
//! 4. **Progress Calculation**
//!    ```vb
//!    percent = (Loc(fileNum) / LOF(fileNum)) * 100
//!    ```
//!
//! 5. **Check Empty File**
//!    ```vb
//!    If LOF(fileNum) = 0 Then
//!        MsgBox "File is empty"
//!    End If
//!    ```
//!
//! 6. **Loop Until End**
//!    ```vb
//!    Do While Loc(fileNum) < LOF(fileNum)
//!        Get #fileNum, , data
//!    Loop
//!    ```
//!
//! 7. **Display File Size**
//!    ```vb
//!    lblSize.Caption = "Size: " & LOF(fileNum) & " bytes"
//!    ```
//!
//! 8. **Allocate Byte Array**
//!    ```vb
//!    ReDim fileData(1 To LOF(fileNum)) As Byte
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Get File Size
//! ```vb
//! Dim fileNum As Integer
//! Dim fileSize As Long
//!
//! fileNum = FreeFile
//! Open "data.bin" For Binary As #fileNum
//!
//! fileSize = LOF(fileNum)
//! MsgBox "File size: " & fileSize & " bytes"
//!
//! Close #fileNum
//! ```
//!
//! ### Example 2: Calculate Record Count
//! ```vb
//! Type CustomerRecord
//!     ID As Long
//!     Name As String * 50
//!     Balance As Currency
//! End Type
//!
//! Dim customer As CustomerRecord
//! Dim fileNum As Integer
//! Dim totalRecords As Long
//!
//! fileNum = FreeFile
//! Open "customers.dat" For Random As #fileNum Len = Len(customer)
//!
//! totalRecords = LOF(fileNum) / Len(customer)
//! MsgBox "Total customers: " & totalRecords
//!
//! Close #fileNum
//! ```
//!
//! ### Example 3: Read Entire File
//! ```vb
//! Dim fileNum As Integer
//! Dim fileContents As String
//! Dim fileSize As Long
//!
//! fileNum = FreeFile
//! Open "readme.txt" For Binary As #fileNum
//!
//! fileSize = LOF(fileNum)
//! fileContents = String(fileSize, 0)
//! Get #fileNum, , fileContents
//!
//! Close #fileNum
//!
//! MsgBox fileContents
//! ```
//!
//! ### Example 4: Progress Indicator
//! ```vb
//! Dim fileNum As Integer
//! Dim data As Byte
//! Dim fileSize As Long
//! Dim bytesRead As Long
//!
//! fileNum = FreeFile
//! Open "large.dat" For Binary As #fileNum
//! fileSize = LOF(fileNum)
//!
//! Do While Loc(fileNum) < fileSize
//!     Get #fileNum, , data
//!     ProcessByte data
//!     
//!     bytesRead = Loc(fileNum)
//!     If bytesRead Mod 1024 = 0 Then
//!         lblProgress.Caption = Format((bytesRead / fileSize) * 100, "0.0") & "%"
//!         DoEvents
//!     End If
//! Loop
//!
//! Close #fileNum
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `GetFileSize`
//! ```vb
//! Function GetFileSize(ByVal fileNum As Integer) As Long
//!     On Error Resume Next
//!     GetFileSize = LOF(fileNum)
//!     If Err.Number <> 0 Then
//!         GetFileSize = -1
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: `CalculateRecordCount`
//! ```vb
//! Function CalculateRecordCount(ByVal fileNum As Integer, _
//!                                ByVal recordLength As Long) As Long
//!     Dim fileSize As Long
//!     fileSize = LOF(fileNum)
//!     
//!     If recordLength > 0 Then
//!         CalculateRecordCount = fileSize \ recordLength
//!     Else
//!         CalculateRecordCount = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `ReadEntireFile`
//! ```vb
//! Function ReadEntireFile(ByVal fileNum As Integer) As String
//!     Dim fileSize As Long
//!     Dim buffer As String
//!     
//!     fileSize = LOF(fileNum)
//!     If fileSize > 0 Then
//!         buffer = String(fileSize, 0)
//!         Get #fileNum, 1, buffer
//!         ReadEntireFile = buffer
//!     Else
//!         ReadEntireFile = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `ReadEntireFileAsBytes`
//! ```vb
//! Function ReadEntireFileAsBytes(ByVal fileNum As Integer) As Byte()
//!     Dim fileSize As Long
//!     Dim buffer() As Byte
//!     
//!     fileSize = LOF(fileNum)
//!     If fileSize > 0 Then
//!         ReDim buffer(0 To fileSize - 1) As Byte
//!         Get #fileNum, 1, buffer
//!         ReadEntireFileAsBytes = buffer
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `CalculateProgress`
//! ```vb
//! Function CalculateProgress(ByVal fileNum As Integer) As Single
//!     Dim currentPos As Long
//!     Dim totalSize As Long
//!     
//!     currentPos = Loc(fileNum)
//!     totalSize = LOF(fileNum)
//!     
//!     If totalSize > 0 Then
//!         CalculateProgress = (currentPos / totalSize) * 100
//!     Else
//!         CalculateProgress = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 6: `IsEmptyFile`
//! ```vb
//! Function IsEmptyFile(ByVal fileNum As Integer) As Boolean
//!     IsEmptyFile = (LOF(fileNum) = 0)
//! End Function
//! ```
//!
//! ### Pattern 7: `GetBytesRemaining`
//! ```vb
//! Function GetBytesRemaining(ByVal fileNum As Integer) As Long
//!     GetBytesRemaining = LOF(fileNum) - Loc(fileNum)
//! End Function
//! ```
//!
//! ### Pattern 8: `FormatFileSize`
//! ```vb
//! Function FormatFileSize(ByVal fileNum As Integer) As String
//!     Dim bytes As Long
//!     bytes = LOF(fileNum)
//!     
//!     If bytes < 1024 Then
//!         FormatFileSize = bytes & " bytes"
//!     ElseIf bytes < 1048576 Then
//!         FormatFileSize = Format(bytes / 1024, "0.0") & " KB"
//!     Else
//!         FormatFileSize = Format(bytes / 1048576, "0.0") & " MB"
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `IsAtEndOfFile`
//! ```vb
//! Function IsAtEndOfFile(ByVal fileNum As Integer) As Boolean
//!     IsAtEndOfFile = (Loc(fileNum) >= LOF(fileNum))
//! End Function
//! ```
//!
//! ### Pattern 10: `AllocateBuffer`
//! ```vb
//! Function AllocateBuffer(ByVal fileNum As Integer) As String
//!     Dim size As Long
//!     size = LOF(fileNum)
//!     
//!     If size > 0 Then
//!         AllocateBuffer = String(size, 0)
//!     Else
//!         AllocateBuffer = ""
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: File Reader Class
//! ```vb
//! ' Class: FileReader
//! Private m_fileNum As Integer
//! Private m_filename As String
//! Private m_fileSize As Long
//! Private m_isOpen As Boolean
//!
//! Public Sub OpenFile(ByVal filename As String)
//!     If m_isOpen Then CloseFile
//!     
//!     m_filename = filename
//!     m_fileNum = FreeFile
//!     Open filename For Binary As #m_fileNum
//!     m_fileSize = LOF(m_fileNum)
//!     m_isOpen = True
//! End Sub
//!
//! Public Property Get Size() As Long
//!     If m_isOpen Then
//!         Size = m_fileSize
//!     Else
//!         Size = 0
//!     End If
//! End Property
//!
//! Public Property Get Position() As Long
//!     If m_isOpen Then
//!         Position = Loc(m_fileNum)
//!     Else
//!         Position = 0
//!     End If
//! End Property
//!
//! Public Property Get Progress() As Single
//!     If m_isOpen And m_fileSize > 0 Then
//!         Progress = (Loc(m_fileNum) / m_fileSize) * 100
//!     Else
//!         Progress = 0
//!     End If
//! End Property
//!
//! Public Property Get IsEOF() As Boolean
//!     If m_isOpen Then
//!         IsEOF = (Loc(m_fileNum) >= m_fileSize)
//!     Else
//!         IsEOF = True
//!     End If
//! End Property
//!
//! Public Property Get IsEmpty() As Boolean
//!     IsEmpty = (m_fileSize = 0)
//! End Property
//!
//! Public Function ReadAll() As String
//!     If m_isOpen And m_fileSize > 0 Then
//!         ReadAll = String(m_fileSize, 0)
//!         Get #m_fileNum, 1, ReadAll
//!     Else
//!         ReadAll = ""
//!     End If
//! End Function
//!
//! Public Sub CloseFile()
//!     If m_isOpen Then
//!         Close #m_fileNum
//!         m_isOpen = False
//!         m_fileSize = 0
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseFile
//! End Sub
//! ```
//!
//! ### Example 2: Random File Manager
//! ```vb
//! ' Class: RandomFileManager
//! Private m_fileNum As Integer
//! Private m_recordLength As Long
//! Private m_totalRecords As Long
//! Private m_isOpen As Boolean
//!
//! Public Sub OpenFile(ByVal filename As String, _
//!                     ByVal recordLength As Long)
//!     If m_isOpen Then CloseFile
//!     
//!     m_recordLength = recordLength
//!     m_fileNum = FreeFile
//!     Open filename For Random As #m_fileNum Len = recordLength
//!     
//!     m_totalRecords = LOF(m_fileNum) \ recordLength
//!     m_isOpen = True
//! End Sub
//!
//! Public Property Get RecordCount() As Long
//!     If m_isOpen Then
//!         RecordCount = m_totalRecords
//!     Else
//!         RecordCount = 0
//!     End If
//! End Property
//!
//! Public Property Get FileSize() As Long
//!     If m_isOpen Then
//!         FileSize = LOF(m_fileNum)
//!     Else
//!         FileSize = 0
//!     End If
//! End Property
//!
//! Public Property Get CurrentRecord() As Long
//!     If m_isOpen Then
//!         CurrentRecord = Loc(m_fileNum)
//!     Else
//!         CurrentRecord = 0
//!     End If
//! End Property
//!
//! Public Function IsValidRecord(ByVal recordNum As Long) As Boolean
//!     IsValidRecord = (recordNum >= 1 And recordNum <= m_totalRecords)
//! End Function
//!
//! Public Sub RefreshRecordCount()
//!     If m_isOpen Then
//!         m_totalRecords = LOF(m_fileNum) \ m_recordLength
//!     End If
//! End Sub
//!
//! Public Sub CloseFile()
//!     If m_isOpen Then
//!         Close #m_fileNum
//!         m_isOpen = False
//!         m_totalRecords = 0
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseFile
//! End Sub
//! ```
//!
//! ### Example 3: File Copy with Progress
//! ```vb
//! Sub CopyFileWithProgress(ByVal sourceFile As String, _
//!                          ByVal destFile As String, _
//!                          Optional ByVal progressBar As ProgressBar = Nothing)
//!     Dim sourceNum As Integer, destNum As Integer
//!     Dim buffer(1 To 4096) As Byte
//!     Dim bytesRead As Long
//!     Dim totalSize As Long
//!     Dim lastPercent As Integer
//!     Dim currentPercent As Integer
//!     
//!     ' Open source file
//!     sourceNum = FreeFile
//!     Open sourceFile For Binary As #sourceNum
//!     totalSize = LOF(sourceNum)
//!     
//!     ' Open destination file
//!     destNum = FreeFile
//!     Open destFile For Binary As #destNum
//!     
//!     ' Copy in chunks
//!     Do While Loc(sourceNum) < totalSize
//!         Get #sourceNum, , buffer
//!         Put #destNum, , buffer
//!         
//!         If Not progressBar Is Nothing Then
//!             bytesRead = Loc(sourceNum)
//!             currentPercent = Int((bytesRead / totalSize) * 100)
//!             
//!             If currentPercent <> lastPercent Then
//!                 progressBar.Value = currentPercent
//!                 lastPercent = currentPercent
//!                 DoEvents
//!             End If
//!         End If
//!     Loop
//!     
//!     Close #sourceNum
//!     Close #destNum
//! End Sub
//! ```
//!
//! ### Example 4: File Information Display
//! ```vb
//! ' Form with labels and progress bar
//! Private m_fileNum As Integer
//! Private m_fileSize As Long
//!
//! Private Sub OpenAndDisplayFile(ByVal filename As String)
//!     m_fileNum = FreeFile
//!     Open filename For Binary As #m_fileNum
//!     m_fileSize = LOF(m_fileNum)
//!     
//!     ' Display file information
//!     lblFilename.Caption = filename
//!     lblFileSize.Caption = FormatBytes(m_fileSize)
//!     ProgressBar1.Min = 0
//!     ProgressBar1.Max = 100
//!     
//!     Timer1.Enabled = True
//! End Sub
//!
//! Private Sub ProcessFile()
//!     Dim data As Byte
//!     
//!     Do While Loc(m_fileNum) < m_fileSize
//!         Get #m_fileNum, , data
//!         ProcessByte data
//!     Loop
//!     
//!     Timer1.Enabled = False
//!     Close #m_fileNum
//!     MsgBox "Processing complete!"
//! End Sub
//!
//! Private Sub Timer1_Timer()
//!     UpdateProgress
//! End Sub
//!
//! Private Sub UpdateProgress()
//!     Dim bytesProcessed As Long
//!     Dim percent As Single
//!     
//!     On Error Resume Next
//!     bytesProcessed = Loc(m_fileNum)
//!     
//!     If m_fileSize > 0 Then
//!         percent = (bytesProcessed / m_fileSize) * 100
//!         ProgressBar1.Value = percent
//!         
//!         lblProgress.Caption = FormatBytes(bytesProcessed) & " of " & _
//!                              FormatBytes(m_fileSize) & " (" & _
//!                              Format(percent, "0.0") & "%)"
//!     End If
//! End Sub
//!
//! Private Function FormatBytes(ByVal bytes As Long) As String
//!     If bytes < 1024 Then
//!         FormatBytes = bytes & " B"
//!     ElseIf bytes < 1048576 Then
//!         FormatBytes = Format(bytes / 1024, "0.0") & " KB"
//!     ElseIf bytes < 1073741824 Then
//!         FormatBytes = Format(bytes / 1048576, "0.0") & " MB"
//!     Else
//!         FormatBytes = Format(bytes / 1073741824, "0.0") & " GB"
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 52: Bad file name or number
//! On Error Resume Next
//! size = LOF(999)
//! If Err.Number = 52 Then
//!     MsgBox "File not open!"
//! End If
//!
//! ' Error 68: Device unavailable
//! size = LOF(fileNum)
//! If Err.Number = 68 Then
//!     MsgBox "Device unavailable!"
//! End If
//!
//! ' Safe size retrieval
//! Function GetSafeFileSize(ByVal fileNum As Integer) As Long
//!     On Error Resume Next
//!     GetSafeFileSize = LOF(fileNum)
//!     If Err.Number <> 0 Then
//!         GetSafeFileSize = -1
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Very Fast**: LOF is a simple file system query
//! - **No I/O**: Does not read file contents
//! - **Cache Result**: Store in variable if using multiple times
//! - **No Side Effects**: Does not change file pointer
//! - **Constant Time**: O(1) operation regardless of file size
//!
//! ## Best Practices
//!
//! 1. **Cache the value** if using LOF multiple times in a loop
//! 2. **Check for zero** to detect empty files
//! 3. **Use for buffer allocation** when reading entire files
//! 4. **Combine with Loc** for progress calculation
//! 5. **Use integer division** (\\) for record count calculation
//! 6. **Handle errors** for unopened files
//! 7. **Check 2GB limit** for very large files
//! 8. **Refresh if writing** as file size may change
//! 9. **Use with Random files** to calculate record count
//! 10. **Prefer over `FileLen`** for open files
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | File State | Return Value |
//! |----------|---------|------------|--------------|
//! | **LOF** | Get file size | Must be open | Size in bytes |
//! | **`FileLen`** | Get file size | Must be closed | Size in bytes |
//! | **Loc** | Get position | Must be open | Current position |
//! | **Seek** | Get/set position | Must be open | Next position |
//! | **EOF** | Check end | Must be open | Boolean |
//!
//! ## LOF vs `FileLen`
//!
//! ```vb
//! ' FileLen - for closed files
//! size = FileLen("data.dat")
//!
//! ' LOF - for open files
//! Open "data.dat" For Binary As #1
//! size = LOF(1)
//! Close #1
//!
//! ' LOF is better for open files:
//! ' - Reflects current size if file is being written
//! ' - Faster (no need to close and reopen)
//! ' - Works with all file modes
//! ```
//!
//! ## Mode-Specific Usage
//!
//! ```vb
//! ' Binary mode - get exact byte count
//! Open "data.bin" For Binary As #1
//! totalBytes = LOF(1)
//!
//! ' Random mode - calculate record count
//! Open "records.dat" For Random As #1 Len = 128
//! totalRecords = LOF(1) \ 128
//!
//! ' Input mode - get file size for progress
//! Open "text.txt" For Input As #1
//! fileSize = LOF(1)
//!
//! ' Output/Append mode - check current size
//! Open "log.txt" For Append As #1
//! currentSize = LOF(1)
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Returns Long (max ~2GB file support)
//! - For files > 2GB, result may overflow (wraps to negative)
//! - Windows-specific file I/O
//! - Behavior identical across Windows versions
//! - Works with local and network files
//! - UNC paths supported
//!
//! ## Limitations
//!
//! - **2GB Limit**: Long type limits to ~2,147,483,647 bytes
//! - **Files > 2GB**: Result overflows and becomes negative
//! - **Requires Open File**: Error 52 if file not open
//! - **No String Files**: Works with file numbers only
//! - **Not for Directories**: Only for files
//! - **Static at Call**: Returns size at moment of call
//! - **No Metadata**: Only returns size, not other attributes
//!
//! ## Related Functions
//!
//! - `Loc`: Get current read/write position in file
//! - `Seek`: Get/set file position
//! - `EOF`: Check if at end of file
//! - `FileLen`: Get length of closed file
//! - `Open`: Open file for I/O
//! - `Close`: Close open file
//! - `FreeFile`: Get available file number
//! - `FileAttr`: Get file mode or handle

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_lof_basic() {
        let source = r#"
            Dim size As Long
            size = LOF(1)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_file_variable() {
        let source = r#"
            fileSize = LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_if_statement() {
        let source = r#"
            If LOF(fileNum) = 0 Then
                MsgBox "File is empty"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_record_count() {
        let source = r#"
            totalRecords = LOF(fileNum) / Len(record)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_do_while() {
        let source = r#"
            Do While Loc(1) < LOF(1)
                Get #1, , data
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_function_return() {
        let source = r#"
            Function GetFileSize() As Long
                GetFileSize = LOF(fileNum)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_string_allocation() {
        let source = r#"
            buffer = String(LOF(fileNum), 0)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_error_handling() {
        let source = r#"
            On Error Resume Next
            size = LOF(fileNum)
            If Err.Number <> 0 Then
                MsgBox "Error getting size"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_progress_calculation() {
        let source = r#"
            percent = (Loc(fileNum) / LOF(fileNum)) * 100
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_label_assignment() {
        let source = r#"
            lblSize.Caption = "Size: " & LOF(fileNum) & " bytes"
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_with_statement() {
        let source = r#"
            With fileInfo
                .Size = LOF(fileNum)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_select_case() {
        let source = r#"
            Select Case LOF(fileNum)
                Case 0
                    MsgBox "Empty"
                Case Else
                    MsgBox "Has data"
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_elseif() {
        let source = r#"
            If LOF(fileNum) = 0 Then
                status = "Empty"
            ElseIf LOF(fileNum) < 1024 Then
                status = "Small"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_parentheses() {
        let source = r#"
            size = (LOF(fileNum))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_iif() {
        let source = r#"
            msg = IIf(LOF(fileNum) > 0, "Has data", "Empty")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_in_class() {
        let source = r#"
            Private Sub Class_Method()
                m_fileSize = LOF(m_fileNum)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_function_argument() {
        let source = r#"
            Call ProcessFileSize(LOF(fileNum))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_property_assignment() {
        let source = r#"
            MyObject.FileSize = LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_array_assignment() {
        let source = r#"
            fileSizes(i) = LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_while_wend() {
        let source = r#"
            While Loc(fileNum) < LOF(fileNum)
                Get #fileNum, , record
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_do_until() {
        let source = r#"
            Do Until Loc(fileNum) >= LOF(fileNum)
                Get #fileNum, , buffer
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_redim() {
        let source = r#"
            ReDim fileData(1 To LOF(fileNum)) As Byte
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_integer_division() {
        let source = r#"
            recordCount = LOF(fileNum) \ recordSize
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_msgbox() {
        let source = r#"
            MsgBox "File size: " & LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_debug_print() {
        let source = r#"
            Debug.Print "Size: " & LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_progressbar() {
        let source = r#"
            ProgressBar1.Max = LOF(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_lof_comparison() {
        let source = r#"
            If LOF(fileNum) > 1048576 Then
                MsgBox "File larger than 1MB"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LOF"));
        assert!(text.contains("Identifier"));
    }
}
