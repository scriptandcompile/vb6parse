//! # EOF Function
//!
//! Returns a Boolean value indicating whether the end of a file opened for Random or sequential
//! Input has been reached.
//!
//! ## Syntax
//!
//! ```vb
//! EOF(filenumber)
//! ```
//!
//! ## Parameters
//!
//! - **filenumber**: Required. An Integer containing a valid file number.
//!
//! ## Return Value
//!
//! Returns a Boolean value. Returns True when the end of a file opened for Random or Input
//! access has been reached; otherwise, returns False.
//!
//! ## Remarks
//!
//! The `EOF` function is used to detect when the end of a file has been reached during
//! sequential or random file reading operations. It's essential for controlling loops that
//! read through files.
//!
//! **Important Characteristics:**
//!
//! - Returns True when end of file is reached
//! - Returns False when more data is available
//! - Works with files opened for Input or Random access
//! - Does not work with files opened for Output or Append
//! - Does not work with binary mode files (use LOF instead)
//! - File must be open before calling EOF
//! - Error if file number is invalid or file is closed
//! - Position-dependent (affected by Get, Input, Line Input)
//! - Can be used to prevent "Input past end of file" error
//!
//! ## File Access Modes
//!
//! - **Input**: Sequential text file reading - EOF returns True after last character
//! - **Random**: Random access files - EOF returns True after last record
//! - **Binary**: Not supported (use LOF and Seek instead)
//! - **Output/Append**: Not applicable (write modes)
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Read text file line by line
//! Dim fileNum As Integer
//! Dim line As String
//!
//! fileNum = FreeFile
//! Open "C:\data.txt" For Input As #fileNum
//!
//! Do Until EOF(fileNum)
//!     Line Input #fileNum, line
//!     Debug.Print line
//! Loop
//!
//! Close #fileNum
//! ```
//!
//! ### Read All Lines into Array
//!
//! ```vb
//! Function ReadAllLines(filePath As String) As Variant
//!     Dim fileNum As Integer
//!     Dim lines() As String
//!     Dim line As String
//!     Dim count As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     count = 0
//!     ReDim lines(0 To 100)
//!     
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         lines(count) = line
//!         count = count + 1
//!         
//!         If count > UBound(lines) Then
//!             ReDim Preserve lines(0 To UBound(lines) + 100)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If count > 0 Then
//!         ReDim Preserve lines(0 To count - 1)
//!         ReadAllLines = lines
//!     Else
//!         ReadAllLines = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Random Access File
//!
//! ```vb
//! Type CustomerRecord
//!     ID As Long
//!     Name As String * 50
//!     Balance As Double
//! End Type
//!
//! Sub ReadAllCustomers()
//!     Dim fileNum As Integer
//!     Dim customer As CustomerRecord
//!     
//!     fileNum = FreeFile
//!     Open "customers.dat" For Random As #fileNum Len = Len(customer)
//!     
//!     Do Until EOF(fileNum)
//!         Get #fileNum, , customer
//!         Debug.Print customer.ID, customer.Name, customer.Balance
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Count Lines in File
//!
//! ```vb
//! Function CountLines(filePath As String) As Long
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim count As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     count = 0
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         count = count + 1
//!     Loop
//!     
//!     Close #fileNum
//!     CountLines = count
//! End Function
//! ```
//!
//! ### Search File for Text
//!
//! ```vb
//! Function FindInFile(filePath As String, searchText As String) As Boolean
//!     Dim fileNum As Integer
//!     Dim line As String
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     FindInFile = False
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         If InStr(1, line, searchText, vbTextCompare) > 0 Then
//!             FindInFile = True
//!             Exit Do
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Function
//! ```
//!
//! ### Read CSV File
//!
//! ```vb
//! Function ReadCSV(filePath As String) As Variant
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim rows() As Variant
//!     Dim rowCount As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     rowCount = 0
//!     ReDim rows(0 To 100)
//!     
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         rows(rowCount) = Split(line, ",")
//!         rowCount = rowCount + 1
//!         
//!         If rowCount > UBound(rows) Then
//!             ReDim Preserve rows(0 To UBound(rows) + 100)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If rowCount > 0 Then
//!         ReDim Preserve rows(0 To rowCount - 1)
//!         ReadCSV = rows
//!     Else
//!         ReadCSV = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Process File with Progress
//!
//! ```vb
//! Sub ProcessFileWithProgress(filePath As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim lineCount As Long
//!     Dim processedCount As Long
//!     
//!     ' Count total lines first
//!     lineCount = CountLines(filePath)
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     processedCount = 0
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         ProcessLine line
//!         processedCount = processedCount + 1
//!         
//!         ' Update progress
//!         If processedCount Mod 100 = 0 Then
//!             lblProgress.Caption = processedCount & " of " & lineCount
//!             DoEvents
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Read Until Marker
//!
//! ```vb
//! Function ReadUntilMarker(fileNum As Integer, marker As String) As String
//!     Dim line As String
//!     Dim content As String
//!     
//!     content = ""
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         If line = marker Then
//!             Exit Do
//!         End If
//!         
//!         content = content & line & vbCrLf
//!     Loop
//!     
//!     ReadUntilMarker = content
//! End Function
//! ```
//!
//! ### Skip Header Lines
//!
//! ```vb
//! Sub ProcessDataFile(filePath As String, headerLines As Integer)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim i As Integer
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     ' Skip header lines
//!     For i = 1 To headerLines
//!         If Not EOF(fileNum) Then
//!             Line Input #fileNum, line
//!         End If
//!     Next i
//!     
//!     ' Process remaining data
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         ProcessDataLine line
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Read Fixed Number of Lines
//!
//! ```vb
//! Function ReadLines(filePath As String, maxLines As Long) As Variant
//!     Dim fileNum As Integer
//!     Dim lines() As String
//!     Dim line As String
//!     Dim count As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     ReDim lines(0 To maxLines - 1)
//!     count = 0
//!     
//!     Do Until EOF(fileNum) Or count >= maxLines
//!         Line Input #fileNum, line
//!         lines(count) = line
//!         count = count + 1
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If count > 0 Then
//!         ReDim Preserve lines(0 To count - 1)
//!         ReadLines = lines
//!     Else
//!         ReadLines = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Merge Multiple Files
//!
//! ```vb
//! Sub MergeFiles(inputFiles As Variant, outputFile As String)
//!     Dim outNum As Integer
//!     Dim inNum As Integer
//!     Dim i As Integer
//!     Dim line As String
//!     
//!     outNum = FreeFile
//!     Open outputFile For Output As #outNum
//!     
//!     For i = LBound(inputFiles) To UBound(inputFiles)
//!         inNum = FreeFile
//!         Open inputFiles(i) For Input As #inNum
//!         
//!         Do Until EOF(inNum)
//!             Line Input #inNum, line
//!             Print #outNum, line
//!         Loop
//!         
//!         Close #inNum
//!     Next i
//!     
//!     Close #outNum
//! End Sub
//! ```
//!
//! ### Filter File Contents
//!
//! ```vb
//! Sub FilterFile(inputFile As String, outputFile As String, filterText As String)
//!     Dim inNum As Integer
//!     Dim outNum As Integer
//!     Dim line As String
//!     
//!     inNum = FreeFile
//!     Open inputFile For Input As #inNum
//!     
//!     outNum = FreeFile
//!     Open outputFile For Output As #outNum
//!     
//!     Do Until EOF(inNum)
//!         Line Input #inNum, line
//!         
//!         If InStr(1, line, filterText, vbTextCompare) > 0 Then
//!             Print #outNum, line
//!         End If
//!     Loop
//!     
//!     Close #inNum
//!     Close #outNum
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Parse Configuration File
//!
//! ```vb
//! Function ParseConfigFile(filePath As String) As Collection
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim config As New Collection
//!     Dim equalPos As Integer
//!     Dim key As String
//!     Dim value As String
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         line = Trim(line)
//!         
//!         ' Skip empty lines and comments
//!         If Len(line) > 0 And Left(line, 1) <> "#" Then
//!             equalPos = InStr(line, "=")
//!             If equalPos > 0 Then
//!                 key = Trim(Left(line, equalPos - 1))
//!                 value = Trim(Mid(line, equalPos + 1))
//!                 
//!                 On Error Resume Next
//!                 config.Add value, key
//!                 On Error GoTo 0
//!             End If
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     Set ParseConfigFile = config
//! End Function
//! ```
//!
//! ### Read Log File with Timestamp Filter
//!
//! ```vb
//! Function ReadLogsSince(logFile As String, sinceDate As Date) As Variant
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim logs() As String
//!     Dim count As Long
//!     Dim lineDate As Date
//!     
//!     fileNum = FreeFile
//!     Open logFile For Input As #fileNum
//!     
//!     count = 0
//!     ReDim logs(0 To 100)
//!     
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         ' Assuming timestamp is first 19 chars: "2025-11-21 10:30:45"
//!         If Len(line) >= 19 Then
//!             On Error Resume Next
//!             lineDate = CDate(Left(line, 19))
//!             
//!             If Err.Number = 0 And lineDate >= sinceDate Then
//!                 logs(count) = line
//!                 count = count + 1
//!                 
//!                 If count > UBound(logs) Then
//!                     ReDim Preserve logs(0 To UBound(logs) + 100)
//!                 End If
//!             End If
//!             On Error GoTo 0
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If count > 0 Then
//!         ReDim Preserve logs(0 To count - 1)
//!         ReadLogsSince = logs
//!     Else
//!         ReadLogsSince = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Batch Process Multiple Files
//!
//! ```vb
//! Sub BatchProcessFiles(filePattern As String)
//!     Dim fileName As String
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim processedCount As Long
//!     
//!     fileName = Dir(filePattern)
//!     Do While fileName <> ""
//!         fileNum = FreeFile
//!         Open fileName For Input As #fileNum
//!         
//!         Do Until EOF(fileNum)
//!             Line Input #fileNum, line
//!             ProcessLine line
//!         Loop
//!         
//!         Close #fileNum
//!         processedCount = processedCount + 1
//!         
//!         fileName = Dir
//!     Loop
//!     
//!     MsgBox processedCount & " files processed"
//! End Sub
//! ```
//!
//! ### Read File in Chunks
//!
//! ```vb
//! Function ReadFileChunk(fileNum As Integer, chunkSize As Long) As Variant
//!     Dim lines() As String
//!     Dim line As String
//!     Dim count As Long
//!     
//!     ReDim lines(0 To chunkSize - 1)
//!     count = 0
//!     
//!     Do Until EOF(fileNum) Or count >= chunkSize
//!         Line Input #fileNum, line
//!         lines(count) = line
//!         count = count + 1
//!     Loop
//!     
//!     If count > 0 Then
//!         ReDim Preserve lines(0 To count - 1)
//!         ReadFileChunk = lines
//!     Else
//!         ReadFileChunk = Array()
//!     End If
//! End Function
//!
//! ' Usage with pagination
//! Sub ProcessLargeFile(filePath As String)
//!     Dim fileNum As Integer
//!     Dim chunk As Variant
//!     Dim pageNum As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     pageNum = 1
//!     Do Until EOF(fileNum)
//!         chunk = ReadFileChunk(fileNum, 1000)
//!         ProcessChunk chunk, pageNum
//!         pageNum = pageNum + 1
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Database Import from Text File
//!
//! ```vb
//! Sub ImportDataFromFile(filePath As String, tableName As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim fields() As String
//!     Dim sql As String
//!     Dim recordCount As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     ' Skip header
//!     If Not EOF(fileNum) Then
//!         Line Input #fileNum, line
//!     End If
//!     
//!     recordCount = 0
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         fields = Split(line, vbTab)
//!         
//!         sql = "INSERT INTO " & tableName & " VALUES ('" & _
//!               Join(fields, "','") & "')"
//!         
//!         ' Execute SQL (pseudo-code)
//!         ExecuteSQL sql
//!         recordCount = recordCount + 1
//!     Loop
//!     
//!     Close #fileNum
//!     MsgBox recordCount & " records imported"
//! End Sub
//! ```
//!
//! ### Create File Index
//!
//! ```vb
//! Type FileIndex
//!     LineNumber As Long
//!     FilePosition As Long
//!     Content As String
//! End Type
//!
//! Function BuildFileIndex(filePath As String) As Variant
//!     Dim fileNum As Integer
//!     Dim index() As FileIndex
//!     Dim line As String
//!     Dim count As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     count = 0
//!     ReDim index(0 To 100)
//!     
//!     Do Until EOF(fileNum)
//!         index(count).LineNumber = count + 1
//!         index(count).FilePosition = Seek(fileNum)
//!         Line Input #fileNum, line
//!         index(count).Content = Left(line, 100)  ' Store first 100 chars
//!         
//!         count = count + 1
//!         If count > UBound(index) Then
//!             ReDim Preserve index(0 To UBound(index) + 100)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If count > 0 Then
//!         ReDim Preserve index(0 To count - 1)
//!         BuildFileIndex = index
//!     Else
//!         BuildFileIndex = Array()
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeReadFile(filePath As String) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileNum As Integer
//!     Dim lines() As String
//!     Dim line As String
//!     Dim count As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     count = 0
//!     ReDim lines(0 To 100)
//!     
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         lines(count) = line
//!         count = count + 1
//!         
//!         If count > UBound(lines) Then
//!             ReDim Preserve lines(0 To UBound(lines) + 100)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     
//!     If count > 0 Then
//!         ReDim Preserve lines(0 To count - 1)
//!         SafeReadFile = lines
//!     Else
//!         SafeReadFile = Array()
//!     End If
//!     Exit Function
//!     
//! ErrorHandler:
//!     If fileNum > 0 Then Close #fileNum
//!     SafeReadFile = Null
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 52** (Bad file name or number): File number is invalid or file is closed
//! - **Error 62** (Input past end of file): Attempted to read beyond EOF without checking
//! - **Error 54** (Bad file mode): File not opened for Input or Random access
//!
//! ## Performance Considerations
//!
//! - `EOF` is very fast (single status check)
//! - No performance penalty for frequent calls
//! - Use `Do Until EOF` rather than counting lines beforehand
//! - For large files, consider buffered reading
//! - Random access files: EOF checks record position
//! - Sequential files: EOF checks character position
//!
//! ## Best Practices
//!
//! ### Always Use EOF to Control File Reading
//!
//! ```vb
//! ' Good - Use EOF to detect end
//! Do Until EOF(fileNum)
//!     Line Input #fileNum, line
//!     ProcessLine line
//! Loop
//!
//! ' Bad - May cause "Input past end of file" error
//! Do While True
//!     Line Input #fileNum, line  ' Error if EOF reached
//!     ProcessLine line
//! Loop
//! ```
//!
//! ### Always Close Files
//!
//! ```vb
//! ' Good - Always close, even on error
//! On Error GoTo ErrorHandler
//! Open filePath For Input As #fileNum
//! Do Until EOF(fileNum)
//!     ' Process
//! Loop
//! Close #fileNum
//! Exit Sub
//!
//! ErrorHandler:
//!     If fileNum > 0 Then Close #fileNum
//! ```
//!
//! ### Check EOF Before Reading
//!
//! ```vb
//! ' Good - Check before reading
//! If Not EOF(fileNum) Then
//!     Line Input #fileNum, line
//! End If
//!
//! ' Or use in loop condition
//! Do Until EOF(fileNum)
//!     Line Input #fileNum, line
//! Loop
//! ```
//!
//! ## Comparison with Other Methods
//!
//! ### EOF vs LOF
//!
//! ```vb
//! ' EOF - Detects end of file for Input/Random
//! Do Until EOF(fileNum)
//!     Line Input #fileNum, line
//! Loop
//!
//! ' LOF - Gets file length (useful for Binary mode)
//! Open file For Binary As #fileNum
//! fileSize = LOF(fileNum)
//! ```
//!
//! ### EOF vs Seek
//!
//! ```vb
//! ' EOF - Boolean end-of-file check
//! isAtEnd = EOF(fileNum)
//!
//! ' Seek - Get/set current position
//! currentPos = Seek(fileNum)
//! ```
//!
//! ## Limitations
//!
//! - Only works with Input and Random access modes
//! - Not applicable to Binary mode (use LOF and Seek)
//! - Not applicable to Output or Append modes
//! - Does not indicate how much data remains
//! - File must be open
//! - Cannot detect EOF before opening file
//!
//! ## Related Functions
//!
//! - `LOF`: Returns length of file in bytes
//! - `Seek`: Returns or sets current position in file
//! - `Open`: Opens file for reading/writing
//! - `Close`: Closes open file
//! - `Line Input`: Reads line from file
//! - `Input`: Reads data from file
//! - `Get`: Reads data from random/binary file
//! - `FreeFile`: Returns next available file number

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn eof_basic() {
        let source = r"
Do Until EOF(1)
    Line Input #1, line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_variable() {
        let source = r"
Do Until EOF(fileNum)
    Line Input #fileNum, line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_in_if_statement() {
        let source = r"
If Not EOF(1) Then
    Line Input #1, line
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_while_loop() {
        let source = r"
Do While Not EOF(fileNum)
    Get #fileNum, , record
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_in_function() {
        let source = r"
Function ReadAllLines(path As String) As Variant
    Do Until EOF(fnum)
        Line Input #fnum, line
    Loop
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_or_condition() {
        let source = r"
Do Until EOF(fileNum) Or count >= maxLines
    Line Input #fileNum, line
    count = count + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_count_lines() {
        let source = r"
count = 0
Do Until EOF(fileNum)
    Line Input #fileNum, line
    count = count + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_exit_do() {
        let source = r#"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    If line = "" Then Exit Do
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_multiple_files() {
        let source = r"
Do Until EOF(inNum)
    Line Input #inNum, line
    Print #outNum, line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_for_loop() {
        let source = r"
For i = 1 To headerLines
    If Not EOF(fileNum) Then
        Line Input #fileNum, line
    End If
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_freefile() {
        let source = r"
fileNum = FreeFile
Open path For Input As #fileNum
Do Until EOF(fileNum)
    Line Input #fileNum, line
Loop
Close #fileNum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_search_file() {
        let source = r"
found = False
Do Until EOF(fileNum)
    Line Input #fileNum, line
    If InStr(line, searchText) > 0 Then
        found = True
        Exit Do
    End If
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_get() {
        let source = r"
Do Until EOF(fileNum)
    Get #fileNum, , customer
    Debug.Print customer.Name
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_csv_reader() {
        let source = r#"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    fields = Split(line, ",")
    ProcessRecord fields
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_doevents() {
        let source = r"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    ProcessLine line
    If lineCount Mod 100 = 0 Then DoEvents
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_error_handling() {
        let source = r"
On Error GoTo ErrorHandler
Do Until EOF(fileNum)
    Line Input #fileNum, line
Loop
Close #fileNum
Exit Sub
ErrorHandler:
If fileNum > 0 Then Close #fileNum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_assignment() {
        let source = r"
isAtEnd = EOF(fileNum)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_nested_loop() {
        let source = r#"
fileName = Dir("*.txt")
Do While fileName <> ""
    fileNum = FreeFile
    Open fileName For Input As #fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, line
    Loop
    Close #fileNum
    fileName = Dir
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_trim() {
        let source = r"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    line = Trim(line)
    If Len(line) > 0 Then ProcessLine line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_debug_print() {
        let source = r"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    Debug.Print line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_array() {
        let source = r"
count = 0
Do Until EOF(fileNum)
    Line Input #fileNum, line
    lines(count) = line
    count = count + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_progress_update() {
        let source = r#"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    processedCount = processedCount + 1
    If processedCount Mod 100 = 0 Then
        lblProgress.Caption = processedCount & " lines"
    End If
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_config_parser() {
        let source = r##"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    If Left(line, 1) <> "#" Then
        ParseConfigLine line
    End If
Loop
"##;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_chunk_reading() {
        let source = r"
Do Until EOF(fileNum)
    chunk = ReadFileChunk(fileNum, 1000)
    ProcessChunk chunk
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn eof_with_seek() {
        let source = r"
Do Until EOF(fileNum)
    position = Seek(fileNum)
    Line Input #fileNum, line
    Debug.Print position, line
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/file/eof");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
