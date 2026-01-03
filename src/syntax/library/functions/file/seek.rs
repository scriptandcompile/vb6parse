/// # Seek Function
///
/// Returns a Long specifying the current read/write position within a file opened using Open statement.
///
/// ## Syntax
///
/// ```vb
/// Seek(filenumber)
/// ```
///
/// ## Parameters
///
/// - `filenumber` - Required. Any valid Integer file number.
///
/// ## Return Value
///
/// Returns a Long value indicating the current position in the file:
/// - For Random mode files: Returns the record number of the next record to be read or written (1-based)
/// - For Binary, Output, Append, and Input mode files: Returns the byte position where the next operation occurs (1-based)
/// - Position 1 is the beginning of the file
///
/// ## Remarks
///
/// The Seek function is used to determine the current position in a file opened with the Open statement.
/// It is particularly useful for:
/// - Determining current position before reading/writing
/// - Saving position to return to later
/// - Calculating file progress (current position vs file size using LOF)
/// - Implementing custom file navigation
/// - Verifying position after Seek statement
///
/// For Random access files, Seek returns the record number (1-based) of the next record to be read or written.
/// For all other access modes (Binary, Input, Output, Append), Seek returns the byte position (1-based).
///
/// The Seek function is different from the Seek statement:
/// - Seek function: Returns current position (read-only operation)
/// - Seek statement: Sets the position for next read/write (write operation)
///
/// The file must be open before calling Seek. Use `FreeFile` to get an available file number.
/// Always close files when done using Close statement.
///
/// ## Typical Uses
///
/// 1. **Current Position**: Determine where you are in the file
/// 2. **Save Position**: Store current position to return to later
/// 3. **Progress Tracking**: Calculate percentage of file processed
/// 4. **Navigation**: Implement forward/backward navigation in files
/// 5. **Record Counting**: Count records processed in Random access
/// 6. **Byte Counting**: Track bytes read/written in Binary mode
/// 7. **Position Verification**: Verify Seek statement worked correctly
/// 8. **File Parsing**: Parse structured binary files with position tracking
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Get current position in binary file
/// Dim FileNum As Integer
/// Dim CurrentPos As Long
/// FileNum = FreeFile
/// Open "data.bin" For Binary As #FileNum
/// CurrentPos = Seek(FileNum)  ' Returns 1 at start
/// Debug.Print "Position: " & CurrentPos
/// Close #FileNum
/// ```
///
/// ```vb
/// ' Example 2: Get current record number in random access
/// Dim FileNum As Integer
/// Dim RecordNum As Long
/// FileNum = FreeFile
/// Open "records.dat" For Random As #FileNum Len = 100
/// RecordNum = Seek(FileNum)  ' Returns next record number
/// Debug.Print "Next Record: " & RecordNum
/// Close #FileNum
/// ```
///
/// ```vb
/// ' Example 3: Save and restore position
/// Dim FileNum As Integer
/// Dim SavedPos As Long
/// FileNum = FreeFile
/// Open "data.txt" For Input As #FileNum
/// ' Read some data...
/// SavedPos = Seek(FileNum)  ' Save position
/// ' Read more data...
/// Seek #FileNum, SavedPos   ' Restore position
/// Close #FileNum
/// ```
///
/// ```vb
/// ' Example 4: Calculate progress percentage
/// Dim FileNum As Integer
/// Dim Progress As Double
/// FileNum = FreeFile
/// Open "large.dat" For Binary As #FileNum
/// Progress = (Seek(FileNum) / LOF(FileNum)) * 100
/// Debug.Print "Progress: " & Format(Progress, "0.0") & "%"
/// Close #FileNum
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: `GetCurrentPosition`
/// Get current file position with error handling
/// ```vb
/// Function GetCurrentPosition(FileNum As Integer) As Long
///     On Error Resume Next
///     GetCurrentPosition = Seek(FileNum)
///     If Err.Number <> 0 Then
///         GetCurrentPosition = -1  ' Error indicator
///     End If
/// End Function
/// ```
///
/// ### Pattern 2: `GetProgressPercentage`
/// Calculate how much of file has been processed
/// ```vb
/// Function GetProgressPercentage(FileNum As Integer) As Double
///     Dim CurrentPos As Long
///     Dim FileSize As Long
///     
///     CurrentPos = Seek(FileNum)
///     FileSize = LOF(FileNum)
///     
///     If FileSize > 0 Then
///         GetProgressPercentage = (CurrentPos / FileSize) * 100
///     Else
///         GetProgressPercentage = 0
///     End If
/// End Function
/// ```
///
/// ### Pattern 3: `IsAtEndOfFile`
/// Check if at end of file using position
/// ```vb
/// Function IsAtEndOfFile(FileNum As Integer) As Boolean
///     IsAtEndOfFile = (Seek(FileNum) > LOF(FileNum))
/// End Function
/// ```
///
/// ### Pattern 4: `GetRemainingBytes`
/// Calculate bytes remaining in file
/// ```vb
/// Function GetRemainingBytes(FileNum As Integer) As Long
///     GetRemainingBytes = LOF(FileNum) - Seek(FileNum) + 1
/// End Function
/// ```
///
/// ### Pattern 5: `SaveAndRestorePosition`
/// Save position, perform operation, restore
/// ```vb
/// Sub SaveAndRestorePosition(FileNum As Integer)
///     Dim SavedPos As Long
///     SavedPos = Seek(FileNum)
///     
///     ' Perform operations that change position
///     Seek #FileNum, 1  ' Go to start
///     ' Read header...
///     
///     Seek #FileNum, SavedPos  ' Restore position
/// End Sub
/// ```
///
/// ### Pattern 6: `GetCurrentRecordNumber`
/// Get current record in Random access file
/// ```vb
/// Function GetCurrentRecordNumber(FileNum As Integer) As Long
///     ' For Random access, Seek returns record number
///     GetCurrentRecordNumber = Seek(FileNum)
/// End Function
/// ```
///
/// ### Pattern 7: `CalculateBytesProcessed`
/// Track bytes processed since start
/// ```vb
/// Function CalculateBytesProcessed(FileNum As Integer, StartPos As Long) As Long
///     CalculateBytesProcessed = Seek(FileNum) - StartPos
/// End Function
/// ```
///
/// ### Pattern 8: `ValidateFilePosition`
/// Verify position is within file bounds
/// ```vb
/// Function ValidateFilePosition(FileNum As Integer) As Boolean
///     Dim CurrentPos As Long
///     Dim FileSize As Long
///     
///     CurrentPos = Seek(FileNum)
///     FileSize = LOF(FileNum)
///     
///     ValidateFilePosition = (CurrentPos >= 1 And CurrentPos <= FileSize + 1)
/// End Function
/// ```
///
/// ### Pattern 9: `GetPositionInfo`
/// Get detailed position information
/// ```vb
/// Sub GetPositionInfo(FileNum As Integer, ByRef Position As Long, _
///                     ByRef Size As Long, ByRef Remaining As Long)
///     Position = Seek(FileNum)
///     Size = LOF(FileNum)
///     Remaining = Size - Position + 1
/// End Sub
/// ```
///
/// ### Pattern 10: `SeekToPercentage`
/// Jump to percentage of file
/// ```vb
/// Sub SeekToPercentage(FileNum As Integer, Percentage As Double)
///     Dim TargetPos As Long
///     TargetPos = CLng((LOF(FileNum) * Percentage) / 100)
///     If TargetPos < 1 Then TargetPos = 1
///     Seek #FileNum, TargetPos
///     Debug.Print "Moved to position: " & Seek(FileNum)
/// End Sub
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: `FilePositionTracker` Class
/// Track and manage file positions with bookmarks
/// ```vb
/// ' Class: FilePositionTracker
/// Private m_FileNum As Integer
/// Private m_Bookmarks As Collection
///
/// Private Sub Class_Initialize()
///     Set m_Bookmarks = New Collection
/// End Sub
///
/// Public Sub Initialize(FileNum As Integer)
///     m_FileNum = FileNum
/// End Sub
///
/// Public Function GetCurrentPosition() As Long
///     GetCurrentPosition = Seek(m_FileNum)
/// End Function
///
/// Public Sub AddBookmark(BookmarkName As String)
///     Dim CurrentPos As Long
///     CurrentPos = Seek(m_FileNum)
///     
///     On Error Resume Next
///     m_Bookmarks.Remove BookmarkName
///     On Error GoTo 0
///     
///     m_Bookmarks.Add CurrentPos, BookmarkName
/// End Sub
///
/// Public Sub GoToBookmark(BookmarkName As String)
///     Dim BookmarkPos As Long
///     On Error Resume Next
///     BookmarkPos = m_Bookmarks(BookmarkName)
///     If Err.Number = 0 Then
///         Seek #m_FileNum, BookmarkPos
///     Else
///         Err.Raise vbObjectError + 1001, "FilePositionTracker", _
///                   "Bookmark not found: " & BookmarkName
///     End If
/// End Sub
///
/// Public Function GetProgress() As Double
///     Dim CurrentPos As Long
///     Dim FileSize As Long
///     
///     CurrentPos = Seek(m_FileNum)
///     FileSize = LOF(m_FileNum)
///     
///     If FileSize > 0 Then
///         GetProgress = (CDbl(CurrentPos) / CDbl(FileSize)) * 100
///     Else
///         GetProgress = 0
///     End If
/// End Function
///
/// Public Function GetRemainingBytes() As Long
///     GetRemainingBytes = LOF(m_FileNum) - Seek(m_FileNum) + 1
/// End Function
///
/// Public Sub ClearBookmarks()
///     Set m_Bookmarks = New Collection
/// End Sub
/// ```
///
/// ### Example 2: `BinaryFileParser` Module
/// Parse structured binary files with position tracking
/// ```vb
/// ' Module: BinaryFileParser
/// Private Type FileHeader
///     Signature As String * 4
///     Version As Integer
///     RecordCount As Long
/// End Type
///
/// Public Function ParseFile(FileName As String) As Collection
///     Dim FileNum As Integer
///     Dim Header As FileHeader
///     Dim Records As New Collection
///     Dim i As Long
///     Dim StartPos As Long
///     
///     FileNum = FreeFile
///     Open FileName For Binary As #FileNum
///     
///     ' Read header
///     StartPos = Seek(FileNum)
///     Debug.Print "Reading header at position: " & StartPos
///     Get #FileNum, , Header
///     
///     Debug.Print "Header read, now at position: " & Seek(FileNum)
///     
///     ' Validate signature
///     If Header.Signature <> "DATA" Then
///         Close #FileNum
///         Err.Raise vbObjectError + 1001, "ParseFile", "Invalid file signature"
///     End If
///     
///     ' Read records
///     For i = 1 To Header.RecordCount
///         Dim RecordPos As Long
///         Dim RecordData As String
///         
///         RecordPos = Seek(FileNum)
///         Debug.Print "Reading record " & i & " at position: " & RecordPos
///         
///         ' Read record (example: 100 byte records)
///         RecordData = Space$(100)
///         Get #FileNum, , RecordData
///         Records.Add RecordData
///     Next i
///     
///     Debug.Print "Finished at position: " & Seek(FileNum)
///     Debug.Print "File size: " & LOF(FileNum)
///     
///     Close #FileNum
///     Set ParseFile = Records
/// End Function
///
/// Public Function GetFileProgress(FileNum As Integer) As String
///     Dim CurrentPos As Long
///     Dim FileSize As Long
///     Dim Percentage As Double
///     
///     CurrentPos = Seek(FileNum)
///     FileSize = LOF(FileNum)
///     
///     If FileSize > 0 Then
///         Percentage = (CDbl(CurrentPos) / CDbl(FileSize)) * 100
///     Else
///         Percentage = 0
///     End If
///     
///     GetFileProgress = "Position: " & CurrentPos & " of " & FileSize & _
///                       " (" & Format(Percentage, "0.0") & "%)"
/// End Function
/// ```
///
/// ### Example 3: `RandomAccessNavigator` Class
/// Navigate through Random access file records
/// ```vb
/// ' Class: RandomAccessNavigator
/// Private m_FileNum As Integer
/// Private m_RecordLength As Integer
/// Private m_TotalRecords As Long
///
/// Public Sub Initialize(FileName As String, RecordLength As Integer)
///     m_RecordLength = RecordLength
///     m_FileNum = FreeFile
///     
///     Open FileName For Random As #m_FileNum Len = RecordLength
///     m_TotalRecords = LOF(m_FileNum) \ RecordLength
/// End Sub
///
/// Public Function GetCurrentRecord() As Long
///     ' For Random access, Seek returns record number
///     GetCurrentRecord = Seek(m_FileNum)
/// End Function
///
/// Public Function GetTotalRecords() As Long
///     GetTotalRecords = m_TotalRecords
/// End Function
///
/// Public Function MoveNext() As Boolean
///     Dim CurrentRecord As Long
///     CurrentRecord = Seek(m_FileNum)
///     
///     If CurrentRecord <= m_TotalRecords Then
///         ' Seek statement will auto-advance after Get
///         MoveNext = True
///     Else
///         MoveNext = False
///     End If
/// End Function
///
/// Public Sub MovePrevious()
///     Dim CurrentRecord As Long
///     CurrentRecord = Seek(m_FileNum)
///     
///     If CurrentRecord > 1 Then
///         Seek #m_FileNum, CurrentRecord - 1
///     End If
/// End Sub
///
/// Public Sub MoveFirst()
///     Seek #m_FileNum, 1
/// End Sub
///
/// Public Sub MoveLast()
///     Seek #m_FileNum, m_TotalRecords
/// End Sub
///
/// Public Function IsAtBeginning() As Boolean
///     IsAtBeginning = (Seek(m_FileNum) = 1)
/// End Function
///
/// Public Function IsAtEnd() As Boolean
///     IsAtEnd = (Seek(m_FileNum) > m_TotalRecords)
/// End Function
///
/// Public Function GetProgress() As Double
///     Dim CurrentRecord As Long
///     CurrentRecord = Seek(m_FileNum)
///     
///     If m_TotalRecords > 0 Then
///         GetProgress = (CDbl(CurrentRecord - 1) / CDbl(m_TotalRecords)) * 100
///     Else
///         GetProgress = 0
///     End If
/// End Function
///
/// Public Sub Close()
///     Close #m_FileNum
/// End Sub
/// ```
///
/// ### Example 4: `FileProgressMonitor` Class
/// Monitor file processing progress with time estimates
/// ```vb
/// ' Class: FileProgressMonitor
/// Private m_FileNum As Integer
/// Private m_StartPosition As Long
/// Private m_StartTime As Double
/// Private m_FileSize As Long
///
/// Public Sub StartMonitoring(FileNum As Integer)
///     m_FileNum = FileNum
///     m_StartPosition = Seek(FileNum)
///     m_StartTime = Timer
///     m_FileSize = LOF(FileNum)
/// End Sub
///
/// Public Function GetCurrentPosition() As Long
///     GetCurrentPosition = Seek(m_FileNum)
/// End Function
///
/// Public Function GetBytesProcessed() As Long
///     GetBytesProcessed = Seek(m_FileNum) - m_StartPosition
/// End Function
///
/// Public Function GetPercentComplete() As Double
///     Dim CurrentPos As Long
///     CurrentPos = Seek(m_FileNum)
///     
///     If m_FileSize > 0 Then
///         GetPercentComplete = (CDbl(CurrentPos) / CDbl(m_FileSize)) * 100
///     Else
///         GetPercentComplete = 0
///     End If
/// End Function
///
/// Public Function GetElapsedSeconds() As Double
///     GetElapsedSeconds = Timer - m_StartTime
/// End Function
///
/// Public Function GetEstimatedTimeRemaining() As Double
///     Dim BytesProcessed As Long
///     Dim BytesRemaining As Long
///     Dim ElapsedTime As Double
///     Dim BytesPerSecond As Double
///     
///     BytesProcessed = Seek(m_FileNum) - m_StartPosition
///     BytesRemaining = m_FileSize - Seek(m_FileNum) + 1
///     ElapsedTime = Timer - m_StartTime
///     
///     If BytesProcessed > 0 And ElapsedTime > 0 Then
///         BytesPerSecond = BytesProcessed / ElapsedTime
///         GetEstimatedTimeRemaining = BytesRemaining / BytesPerSecond
///     Else
///         GetEstimatedTimeRemaining = 0
///     End If
/// End Function
///
/// Public Function GetProgressReport() As String
///     Dim Report As String
///     Dim PercentComplete As Double
///     Dim ElapsedTime As Double
///     Dim RemainingTime As Double
///     
///     PercentComplete = GetPercentComplete()
///     ElapsedTime = GetElapsedSeconds()
///     RemainingTime = GetEstimatedTimeRemaining()
///     
///     Report = "Progress: " & Format(PercentComplete, "0.0") & "%" & vbCrLf
///     Report = Report & "Position: " & Seek(m_FileNum) & " of " & m_FileSize & vbCrLf
///     Report = Report & "Elapsed: " & Format(ElapsedTime, "0.0") & "s" & vbCrLf
///     Report = Report & "Remaining: " & Format(RemainingTime, "0.0") & "s"
///     
///     GetProgressReport = Report
/// End Function
///
/// Public Sub Reset()
///     m_StartPosition = Seek(m_FileNum)
///     m_StartTime = Timer
/// End Sub
/// ```
///
/// ## Error Handling
///
/// The Seek function can generate the following errors:
///
/// - **Error 52** (Bad file name or number): File not open or invalid file number
/// - **Error 5** (Invalid procedure call): File number is invalid
///
/// Always use error handling when working with file I/O:
/// ```vb
/// On Error Resume Next
/// CurrentPos = Seek(FileNum)
/// If Err.Number <> 0 Then
///     MsgBox "Error getting position: " & Err.Description
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - Seek function is very fast (just reads internal file pointer)
/// - No disk I/O involved (unlike reading/writing data)
/// - Can be called frequently without performance impact
/// - Useful for progress tracking in loops without overhead
/// - Combine with LOF for efficient progress calculations
///
/// ## Best Practices
///
/// 1. **Validate File Number**: Ensure file is open before calling Seek
/// 2. **Use with LOF**: Combine with LOF function to calculate progress
/// 3. **Error Handling**: Always use error handling for file operations
/// 4. **Close Files**: Always close files when done to free resources
/// 5. **Save Position**: Store position before operations that change it
/// 6. **1-Based Position**: Remember positions start at 1, not 0
/// 7. **Mode Awareness**: Know whether function returns record# or byte position
/// 8. **`FreeFile` Usage**: Use `FreeFile` to get available file numbers
/// 9. **Progress Updates**: Update UI periodically, not on every byte
/// 10. **Position Validation**: Verify position is within expected range
///
/// ## Comparison with Related Functions
///
/// | Function/Statement | Purpose | Returns | Mode-Specific |
/// |-------------------|---------|---------|---------------|
/// | Seek (Function) | Get current position | Long (position) | Record# or byte position |
/// | Seek (Statement) | Set position | N/A (void) | Sets for next operation |
/// | LOF | Get file size | Long (bytes) | Total file size |
/// | EOF | Check end of file | Boolean | True if past end |
/// | Loc | Get current position | Long | Similar to Seek |
/// | `FileLen` | Get file length | Long (bytes) | File must be closed |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - 1-based positioning (unlike many other languages)
/// - Maximum file size limited by Long data type (2GB)
/// - For files > 2GB, position may overflow
/// - Random access mode returns record numbers
/// - All other modes return byte positions
///
/// ## Limitations
///
/// - Cannot handle files larger than 2GB (Long overflow)
/// - File must be open before calling Seek
/// - Returns position for next operation (not last operation)
/// - For Random access, assumes fixed-length records
/// - No support for 64-bit file positions
/// - Cannot determine position in closed files
///
/// ## Related Functions
///
/// - `Seek` (Statement): Set the position for next read/write operation
/// - `LOF`: Get the length of an open file in bytes
/// - `EOF`: Determine if end of file has been reached
/// - `Loc`: Get the current read/write position (similar to Seek function)
/// - `Get`: Read data from file (advances position)
/// - `Put`: Write data to file (advances position)
/// - `Open`: Open file for I/O operations
/// - `Close`: Close an open file
/// - `FreeFile`: Get next available file number
///
#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn seek_basic() {
        let source = r"
Sub Test()
    Dim pos As Long
    pos = Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("pos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("pos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_with_variable() {
        let source = r"
Sub Test()
    Dim fileNum As Integer
    Dim position As Long
    fileNum = 1
    position = Seek(fileNum)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("fileNum"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("position"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileNum"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("position"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileNum"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_if_statement() {
        let source = r#"
Sub Test()
    If Seek(1) > 100 Then
        MsgBox "Past position 100"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Past position 100\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_function_return() {
        let source = r"
Function GetPosition() As Long
    GetPosition = Seek(1)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetPosition"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetPosition"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_variable_assignment() {
        let source = r"
Sub Test()
    Dim currentPos As Long
    currentPos = Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("currentPos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("currentPos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Position: " & Seek(1)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Position: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        SeekKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_debug_print() {
        let source = r"
Sub Test()
    Debug.Print Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        SeekKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_select_case() {
        let source = r#"
Sub Test()
    Select Case Seek(1)
        Case 1
            MsgBox "At start"
        Case Else
            MsgBox "Not at start"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"At start\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Not at start\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_class_usage() {
        let source = r"
Class FileManager
    Public Function GetCurrentPosition(fileNum As Integer) As Long
        GetCurrentPosition = Seek(fileNum)
    End Function
End Class
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("FileManager"),
                Newline,
            },
            Whitespace,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetCurrentPosition"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fileNum"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetCurrentPosition"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileNum"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
        ]);
    }

    #[test]
    fn seek_with_statement() {
        let source = r"
Sub Test()
    With FileManager
        Dim pos As Long
        pos = Seek(1)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("FileManager"),
                        Newline,
                        StatementList {
                            Whitespace,
                            DimStatement {
                                DimKeyword,
                                Whitespace,
                                Identifier ("pos"),
                                Whitespace,
                                AsKeyword,
                                Whitespace,
                                LongKeyword,
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("pos"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    SeekKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_elseif() {
        let source = r#"
Sub Test()
    If Seek(1) = 1 Then
        MsgBox "At start"
    ElseIf Seek(1) > 100 Then
        MsgBox "Past 100"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"At start\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseIfClause {
                            ElseIfKeyword,
                            Whitespace,
                            BinaryExpression {
                                CallExpression {
                                    SeekKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Whitespace,
                            ThenKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Past 100\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_for_loop() {
        let source = r"
Sub Test()
    Dim i As Long
    For i = 1 To 10
        Debug.Print Seek(1)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                SeekKeyword,
                                LeftParenthesis,
                                IntegerLiteral ("1"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_do_while() {
        let source = r"
Sub Test()
    Do While Seek(1) < 1000
        ' Read data
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1000"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_do_until() {
        let source = r"
Sub Test()
    Do Until Seek(1) > LOF(1)
        ' Read data
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("LOF"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_while_wend() {
        let source = r"
Sub Test()
    While Seek(1) < 1000
        ' Read data
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1000"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_parentheses() {
        let source = r"
Sub Test()
    Dim result As Long
    result = (Seek(1) + 100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    SeekKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_iif() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = IIf(Seek(1) = 1, "Start", "Not start")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("msg"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        CallExpression {
                                            SeekKeyword,
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("1"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Start\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Not start\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_array_assignment() {
        let source = r"
Sub Test()
    Dim positions(10) As Long
    positions(0) = Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("positions"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("positions"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_property_assignment() {
        let source = r"
Class FileInfo
    Public Position As Long
End Class

Sub Test()
    Dim info As New FileInfo
    info.Position = Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("FileInfo"),
                Newline,
            },
            Whitespace,
            DimStatement {
                PublicKeyword,
                Whitespace,
                Identifier ("Position"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("info"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("FileInfo"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("info"),
                            PeriodOperator,
                            Identifier ("Position"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_function_argument() {
        let source = r"
Sub ProcessPosition(pos As Long)
End Sub

Sub Test()
    ProcessPosition Seek(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessPosition"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("pos"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessPosition"),
                        Whitespace,
                        SeekKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_concatenation() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = "Current position: " & Seek(1)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("msg"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Current position: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                SeekKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_comparison() {
        let source = r"
Sub Test()
    Dim atEnd As Boolean
    atEnd = (Seek(1) > LOF(1))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("atEnd"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("atEnd"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    SeekKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("LOF"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_arithmetic() {
        let source = r"
Sub Test()
    Dim remaining As Long
    remaining = LOF(1) - Seek(1) + 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        RemComment,
                        Newline,
                    },
                    Whitespace,
                    RemComment,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_with_lof() {
        let source = r"
Sub Test()
    Dim progress As Double
    progress = (Seek(1) / LOF(1)) * 100
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("progress"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("progress"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    CallExpression {
                                        SeekKeyword,
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                    Whitespace,
                                    DivisionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("LOF"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            MultiplicationOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_freefile() {
        let source = r"
Sub Test()
    Dim fileNum As Integer
    Dim pos As Long
    fileNum = FreeFile
    pos = Seek(fileNum)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("fileNum"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("pos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileNum"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("FreeFile"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("pos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileNum"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim pos As Long
    pos = Seek(1)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("pos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("pos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Error\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn seek_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim position As Long
    position = Seek(1)
    Exit Sub
ErrorHandler:
    MsgBox "Error getting position"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("position"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("position"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            SeekKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error getting position\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
