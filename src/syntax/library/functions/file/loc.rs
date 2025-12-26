//! # Loc Function
//!
//! Returns a Long specifying the current read/write position within an open file.
//!
//! ## Syntax
//!
//! ```vb
//! Loc(filenumber)
//! ```
//!
//! ## Parameters
//!
//! - `filenumber` (Required): Integer file number used in the Open statement
//!   - Must be a valid file number from a currently open file
//!   - File numbers typically obtained from `FreeFile` function
//!
//! ## Return Value
//!
//! Returns a Long:
//! - For Random mode: Record number of last record read or written
//! - For Sequential mode: Current byte position divided by 128
//! - For Binary mode: Position of last byte read or written
//! - Returns 0 if no read/write operations have occurred yet
//! - Returns value based on last I/O operation
//!
//! ## Remarks
//!
//! The Loc function returns the current position in an open file:
//!
//! - Behavior varies based on file access mode
//! - For Random access: Returns record number (1-based)
//! - For Sequential access: Returns byte position / 128 (approximation)
//! - For Binary access: Returns byte position (0-based)
//! - Does not move the file pointer
//! - Read-only operation (non-destructive)
//! - Useful for tracking progress in file operations
//! - Returns position of last operation, not next operation
//! - For Random files, increments after Get/Put
//! - For Binary files, tracks exact byte position
//! - For Sequential files, provides approximate position
//! - Essential for file I/O progress tracking
//! - Used with Seek to navigate files
//! - Different from Seek function (which also sets position)
//! - LOF function returns file length, Loc returns position
//! - Error 52 if file number not open
//! - Error 68 if device unavailable
//! - Common in loops reading/writing files
//! - Helps detect end-of-file conditions
//! - Used for progress bars during file operations
//!
//! ## Typical Uses
//!
//! 1. **Track Random File Position**
//!    ```vb
//!    currentRecord = Loc(1)
//!    ```
//!
//! 2. **Track Binary File Position**
//!    ```vb
//!    bytesProcessed = Loc(fileNum)
//!    ```
//!
//! 3. **Progress Calculation**
//!    ```vb
//!    percentComplete = (Loc(1) / LOF(1)) * 100
//!    ```
//!
//! 4. **Check if Data Written**
//!    ```vb
//!    If Loc(fileNum) > 0 Then
//!        ' File has been written to
//!    End If
//!    ```
//!
//! 5. **Loop Until End**
//!    ```vb
//!    Do While Loc(1) < LOF(1)
//!        Get #1, , record
//!    Loop
//!    ```
//!
//! 6. **Record Number Display**
//!    ```vb
//!    lblRecordNum.Caption = "Record: " & Loc(1)
//!    ```
//!
//! 7. **Byte Position Check**
//!    ```vb
//!    Debug.Print "Position: " & Loc(fileNum)
//!    ```
//!
//! 8. **Progress Bar Update**
//!    ```vb
//!    ProgressBar1.Value = (Loc(1) / totalRecords) * 100
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Random File Position
//! ```vb
//! Type CustomerRecord
//!     ID As Long
//!     Name As String * 50
//! End Type
//!
//! Dim customer As CustomerRecord
//! Dim fileNum As Integer
//!
//! fileNum = FreeFile
//! Open "customers.dat" For Random As #fileNum Len = Len(customer)
//!
//! ' Read records
//! Do While Not EOF(fileNum)
//!     Get #fileNum, , customer
//!     Debug.Print "Record: " & Loc(fileNum)
//! Loop
//!
//! Close #fileNum
//! ```
//!
//! ### Example 2: Binary File Progress
//! ```vb
//! Dim fileNum As Integer
//! Dim data As Byte
//! Dim fileSize As Long
//!
//! fileNum = FreeFile
//! Open "data.bin" For Binary As #fileNum
//! fileSize = LOF(fileNum)
//!
//! Do While Loc(fileNum) < fileSize
//!     Get #fileNum, , data
//!     
//!     ' Update progress
//!     If Loc(fileNum) Mod 1024 = 0 Then
//!         Debug.Print "Progress: " & (Loc(fileNum) / fileSize) * 100 & "%"
//!     End If
//! Loop
//!
//! Close #fileNum
//! ```
//!
//! ### Example 3: Sequential File Position
//! ```vb
//! Dim fileNum As Integer
//! Dim line As String
//!
//! fileNum = FreeFile
//! Open "log.txt" For Input As #fileNum
//!
//! Do While Not EOF(fileNum)
//!     Line Input #fileNum, line
//!     
//!     ' Approximate position (bytes / 128)
//!     Debug.Print "Position: " & Loc(fileNum)
//! Loop
//!
//! Close #fileNum
//! ```
//!
//! ### Example 4: Track Write Position
//! ```vb
//! Dim fileNum As Integer
//! Dim i As Integer
//!
//! fileNum = FreeFile
//! Open "output.bin" For Binary As #fileNum
//!
//! For i = 1 To 100
//!     Put #fileNum, , i
//!     Debug.Print "Wrote to position: " & Loc(fileNum)
//! Next i
//!
//! Close #fileNum
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `CalculateProgress`
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
//! ### Pattern 2: `IsFilePositionChanged`
//! ```vb
//! Function IsFilePositionChanged(ByVal fileNum As Integer, _
//!                                 ByVal lastPosition As Long) As Boolean
//!     IsFilePositionChanged = (Loc(fileNum) <> lastPosition)
//! End Function
//! ```
//!
//! ### Pattern 3: `GetCurrentRecord`
//! ```vb
//! Function GetCurrentRecord(ByVal fileNum As Integer) As Long
//!     ' For Random access files
//!     GetCurrentRecord = Loc(fileNum)
//! End Function
//! ```
//!
//! ### Pattern 4: `GetBytesProcessed`
//! ```vb
//! Function GetBytesProcessed(ByVal fileNum As Integer) As Long
//!     ' For Binary access files
//!     GetBytesProcessed = Loc(fileNum)
//! End Function
//! ```
//!
//! ### Pattern 5: `UpdateProgressBar`
//! ```vb
//! Sub UpdateProgressBar(ByVal fileNum As Integer, _
//!                       ByVal progressBar As ProgressBar)
//!     Dim percent As Single
//!     percent = (Loc(fileNum) / LOF(fileNum)) * 100
//!     
//!     If percent <= 100 Then
//!         progressBar.Value = percent
//!     End If
//!     DoEvents
//! End Sub
//! ```
//!
//! ### Pattern 6: `ReadFileWithProgress`
//! ```vb
//! Sub ReadFileWithProgress(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim data As Byte
//!     Dim lastPercent As Integer
//!     Dim currentPercent As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     
//!     Do While Loc(fileNum) < LOF(fileNum)
//!         Get #fileNum, , data
//!         ProcessByte data
//!         
//!         currentPercent = Int((Loc(fileNum) / LOF(fileNum)) * 100)
//!         If currentPercent <> lastPercent Then
//!             Debug.Print "Progress: " & currentPercent & "%"
//!             lastPercent = currentPercent
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Pattern 7: `GetRecordPosition`
//! ```vb
//! Function GetRecordPosition(ByVal fileNum As Integer) As String
//!     Dim current As Long
//!     Dim total As Long
//!     
//!     current = Loc(fileNum)
//!     total = LOF(fileNum) / Len(recordVariable)
//!     
//!     GetRecordPosition = current & " of " & total
//! End Function
//! ```
//!
//! ### Pattern 8: `SafeLoc`
//! ```vb
//! Function SafeLoc(ByVal fileNum As Integer) As Long
//!     On Error Resume Next
//!     SafeLoc = Loc(fileNum)
//!     If Err.Number <> 0 Then
//!         SafeLoc = -1
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `IsAtEndOfFile`
//! ```vb
//! Function IsAtEndOfFile(ByVal fileNum As Integer) As Boolean
//!     ' For Binary mode
//!     IsAtEndOfFile = (Loc(fileNum) >= LOF(fileNum))
//! End Function
//! ```
//!
//! ### Pattern 10: `LogFilePosition`
//! ```vb
//! Sub LogFilePosition(ByVal fileNum As Integer, _
//!                     ByVal operation As String)
//!     Debug.Print operation & " - Position: " & Loc(fileNum) & _
//!                 " of " & LOF(fileNum)
//! End Sub
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: File Reader with Progress
//! ```vb
//! ' Class: BinaryFileReader
//! Private m_fileNum As Integer
//! Private m_filename As String
//! Private m_fileSize As Long
//!
//! Public Sub OpenFile(ByVal filename As String)
//!     m_filename = filename
//!     m_fileNum = FreeFile
//!     Open filename For Binary As #m_fileNum
//!     m_fileSize = LOF(m_fileNum)
//! End Sub
//!
//! Public Function ReadByte() As Byte
//!     If Not IsEOF Then
//!         Get #m_fileNum, , ReadByte
//!     End If
//! End Function
//!
//! Public Property Get Position() As Long
//!     Position = Loc(m_fileNum)
//! End Property
//!
//! Public Property Get Size() As Long
//!     Size = m_fileSize
//! End Property
//!
//! Public Property Get Progress() As Single
//!     If m_fileSize > 0 Then
//!         Progress = (Loc(m_fileNum) / m_fileSize) * 100
//!     Else
//!         Progress = 0
//!     End If
//! End Property
//!
//! Public Property Get IsEOF() As Boolean
//!     IsEOF = (Loc(m_fileNum) >= m_fileSize)
//! End Property
//!
//! Public Sub CloseFile()
//!     If m_fileNum > 0 Then
//!         Close #m_fileNum
//!         m_fileNum = 0
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseFile
//! End Sub
//! ```
//!
//! ### Example 2: Random File Navigator
//! ```vb
//! ' Class: RandomFileNavigator
//! Private m_fileNum As Integer
//! Private m_recordLength As Integer
//! Private m_totalRecords As Long
//!
//! Public Sub OpenFile(ByVal filename As String, _
//!                     ByVal recordLength As Integer)
//!     m_recordLength = recordLength
//!     m_fileNum = FreeFile
//!     Open filename For Random As #m_fileNum Len = recordLength
//!     m_totalRecords = LOF(m_fileNum) / recordLength
//! End Sub
//!
//! Public Property Get CurrentRecord() As Long
//!     CurrentRecord = Loc(m_fileNum)
//! End Property
//!
//! Public Property Get TotalRecords() As Long
//!     TotalRecords = m_totalRecords
//! End Property
//!
//! Public Property Get ProgressPercent() As Single
//!     If m_totalRecords > 0 Then
//!         ProgressPercent = (Loc(m_fileNum) / m_totalRecords) * 100
//!     Else
//!         ProgressPercent = 0
//!     End If
//! End Property
//!
//! Public Function IsFirstRecord() As Boolean
//!     IsFirstRecord = (Loc(m_fileNum) = 1)
//! End Function
//!
//! Public Function IsLastRecord() As Boolean
//!     IsLastRecord = (Loc(m_fileNum) = m_totalRecords)
//! End Function
//!
//! Public Sub CloseFile()
//!     If m_fileNum > 0 Then
//!         Close #m_fileNum
//!         m_fileNum = 0
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
//!                          Optional progressBar As ProgressBar = Nothing)
//!     Dim sourceNum As Integer, destNum As Integer
//!     Dim buffer As Byte
//!     Dim totalSize As Long
//!     Dim lastPercent As Integer
//!     Dim currentPercent As Integer
//!     
//!     sourceNum = FreeFile
//!     Open sourceFile For Binary As #sourceNum
//!     totalSize = LOF(sourceNum)
//!     
//!     destNum = FreeFile
//!     Open destFile For Binary As #destNum
//!     
//!     Do While Loc(sourceNum) < totalSize
//!         Get #sourceNum, , buffer
//!         Put #destNum, , buffer
//!         
//!         If Not progressBar Is Nothing Then
//!             currentPercent = Int((Loc(sourceNum) / totalSize) * 100)
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
//! ### Example 4: File Processing Monitor
//! ```vb
//! ' Form with lblStatus, lblProgress, ProgressBar1
//! Private m_fileNum As Integer
//! Private m_totalSize As Long
//!
//! Private Sub ProcessLargeFile(ByVal filename As String)
//!     Dim data As Byte
//!     Dim startTime As Single
//!     
//!     m_fileNum = FreeFile
//!     Open filename For Binary As #m_fileNum
//!     m_totalSize = LOF(m_fileNum)
//!     
//!     startTime = Timer
//!     Timer1.Enabled = True
//!     
//!     Do While Loc(m_fileNum) < m_totalSize
//!         Get #m_fileNum, , data
//!         ProcessData data
//!     Loop
//!     
//!     Timer1.Enabled = False
//!     Close #m_fileNum
//!     
//!     lblStatus.Caption = "Complete!"
//! End Sub
//!
//! Private Sub Timer1_Timer()
//!     UpdateProgress
//! End Sub
//!
//! Private Sub UpdateProgress()
//!     Dim percent As Single
//!     Dim bytesProcessed As Long
//!     
//!     On Error Resume Next
//!     bytesProcessed = Loc(m_fileNum)
//!     
//!     If m_totalSize > 0 Then
//!         percent = (bytesProcessed / m_totalSize) * 100
//!         ProgressBar1.Value = percent
//!         lblProgress.Caption = Format(percent, "0.0") & "% - " & _
//!                              FormatBytes(bytesProcessed) & " of " & _
//!                              FormatBytes(m_totalSize)
//!     End If
//! End Sub
//!
//! Private Function FormatBytes(ByVal bytes As Long) As String
//!     If bytes < 1024 Then
//!         FormatBytes = bytes & " bytes"
//!     ElseIf bytes < 1048576 Then
//!         FormatBytes = Format(bytes / 1024, "0.0") & " KB"
//!     Else
//!         FormatBytes = Format(bytes / 1048576, "0.0") & " MB"
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 52: Bad file name or number
//! On Error Resume Next
//! pos = Loc(999)
//! If Err.Number = 52 Then
//!     MsgBox "File not open!"
//! End If
//!
//! ' Error 68: Device unavailable
//! pos = Loc(fileNum)
//! If Err.Number = 68 Then
//!     MsgBox "Device unavailable!"
//! End If
//!
//! ' Safe position retrieval
//! Function GetSafePosition(ByVal fileNum As Integer) As Long
//!     On Error Resume Next
//!     GetSafePosition = Loc(fileNum)
//!     If Err.Number <> 0 Then
//!         GetSafePosition = -1
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Very Fast**: Loc is a simple file pointer query
//! - **No I/O**: Does not perform actual file operations
//! - **Frequent Calls**: Safe to call in tight loops
//! - **Progress Updates**: Use modulo to update UI less frequently
//! - **`DoEvents`**: Call `DoEvents` when updating UI to maintain responsiveness
//!
//! ## Best Practices
//!
//! 1. **Use with LOF** for calculating percentage complete
//! 2. **Check file is open** before calling Loc
//! 3. **Update progress periodically** not on every byte
//! 4. **Cache in variable** if using multiple times
//! 5. **Use for Binary/Random** files (Sequential returns approximation)
//! 6. **Combine with EOF** for robust loop conditions
//! 7. **Handle errors** for unopened files
//! 8. **Use `DoEvents`** when updating UI in loops
//! 9. **Consider mode** when interpreting return value
//! 10. **Document units** (bytes, records, or approximation)
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Read/Write | Return Value |
//! |----------|---------|------------|--------------|
//! | **Loc** | Get position | Read-only | Position |
//! | **Seek** (function) | Get position | Read-only | Position + 1 |
//! | **Seek** (statement) | Set position | Write | N/A |
//! | **LOF** | Get file length | Read-only | Total bytes |
//! | **EOF** | Check end | Read-only | Boolean |
//!
//! ## Loc vs Seek Function
//!
//! ```vb
//! ' Loc - returns position of last operation
//! currentPos = Loc(fileNum)
//!
//! ' Seek - returns next read/write position (Loc + 1)
//! nextPos = Seek(fileNum)
//!
//! ' For Binary mode:
//! ' After reading byte at position 100:
//! ' Loc returns 100
//! ' Seek returns 101
//! ```
//!
//! ## Mode-Specific Behavior
//!
//! ```vb
//! ' Random mode - record number
//! Open "data.dat" For Random As #1 Len = 128
//! Get #1, 5, record
//! Debug.Print Loc(1)  ' Returns 5 (record number)
//!
//! ' Binary mode - byte position
//! Open "data.bin" For Binary As #1
//! Get #1, , buffer
//! Debug.Print Loc(1)  ' Returns bytes read
//!
//! ' Sequential mode - approximate (bytes / 128)
//! Open "text.txt" For Input As #1
//! Line Input #1, line
//! Debug.Print Loc(1)  ' Returns approximate position
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Returns Long (max 2GB file support)
//! - For files > 2GB, result may overflow
//! - Windows-specific file I/O
//! - Behavior identical across Windows versions
//! - Sequential mode approximation may vary
//! - Random mode most reliable
//! - Binary mode exact for files < 2GB
//!
//! ## Limitations
//!
//! - **2GB Limit**: Long type limits file size to ~2GB
//! - **Sequential Approximation**: Not exact for Input/Output/Append modes
//! - **Division by 128**: Sequential mode uses this approximation
//! - **No String Files**: Works with file numbers only
//! - **Requires Open File**: Error if file not open
//! - **Mode Dependent**: Return value meaning varies by mode
//! - **No Directory**: Only for files, not directories
//! - **Last Operation**: Returns position of last I/O, not current
//!
//! ## Related Functions
//!
//! - `Seek`: Get/set file position (next position, not last)
//! - `LOF`: Get length of file
//! - `EOF`: Check if at end of file
//! - `Open`: Open file for I/O
//! - `Close`: Close file
//! - `Get`: Read from file (updates Loc)
//! - `Put`: Write to file (updates Loc)
//! - `FreeFile`: Get available file number

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn loc_basic() {
        let source = r"
            Dim pos As Long
            pos = Loc(1)
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_file_variable() {
        let source = r"
            currentPos = Loc(fileNum)
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_if_statement() {
        let source = r#"
            If Loc(fileNum) > 0 Then
                MsgBox "Data written"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_progress_calculation() {
        let source = r"
            percentComplete = (Loc(1) / LOF(1)) * 100
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_do_while() {
        let source = r"
            Do While Loc(1) < LOF(1)
                Get #1, , data
            Loop
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_function_return() {
        let source = r"
            Function GetPosition() As Long
                GetPosition = Loc(fileNum)
            End Function
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_debug_print() {
        let source = r#"
            Debug.Print "Position: " & Loc(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_error_handling() {
        let source = r#"
            On Error Resume Next
            pos = Loc(fileNum)
            If Err.Number <> 0 Then
                MsgBox "Error reading position"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_comparison() {
        let source = r#"
            If Loc(fileNum) >= LOF(fileNum) Then
                MsgBox "At end of file"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_label_assignment() {
        let source = r#"
            lblPosition.Caption = "Record: " & Loc(1)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_with_statement() {
        let source = r"
            With fileInfo
                .Position = Loc(fileNum)
            End With
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_select_case() {
        let source = r#"
            Select Case Loc(fileNum)
                Case 0
                    MsgBox "No data"
                Case Else
                    MsgBox "Processing"
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_elseif() {
        let source = r#"
            If Loc(fileNum) = 0 Then
                status = "Start"
            ElseIf Loc(fileNum) < 100 Then
                status = "Progress"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_parentheses() {
        let source = r"
            pos = (Loc(fileNum))
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_iif() {
        let source = r#"
            msg = IIf(Loc(fileNum) > 0, "Data exists", "Empty")
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_in_class() {
        let source = r"
            Private Sub Class_Method()
                m_position = Loc(m_fileNum)
            End Sub
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_function_argument() {
        let source = r"
            Call UpdateProgress(Loc(fileNum))
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_property_assignment() {
        let source = r"
            MyObject.Position = Loc(fileNum)
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_array_assignment() {
        let source = r"
            positions(i) = Loc(fileNum)
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_while_wend() {
        let source = r"
            While Loc(fileNum) < totalRecords
                Get #fileNum, , record
            Wend
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_do_until() {
        let source = r"
            Do Until Loc(fileNum) >= targetPosition
                Get #fileNum, , buffer
            Loop
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_for_loop() {
        let source = r"
            For i = 1 To Loc(fileNum)
                ProcessRecord i
            Next i
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_modulo() {
        let source = r"
            If Loc(fileNum) Mod 1024 = 0 Then
                UpdateProgress
            End If
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_progressbar() {
        let source = r"
            ProgressBar1.Value = (Loc(1) / LOF(1)) * 100
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_msgbox() {
        let source = r#"
            MsgBox "Current position: " & Loc(fileNum)
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_format() {
        let source = r#"
            lblStatus.Caption = Format(Loc(fileNum), "0,000")
        "#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn loc_arithmetic() {
        let source = r"
            bytesRemaining = LOF(fileNum) - Loc(fileNum)
        ";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Loc"));
        assert!(text.contains("Identifier"));
    }
}
