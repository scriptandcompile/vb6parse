//! # `DoEvents` Function
//!
//! Yields execution so that the operating system can process other events and messages.
//!
//! ## Syntax
//!
//! ```vb
//! DoEvents()
//! ```
//!
//! ## Parameters
//!
//! None.
//!
//! ## Return Value
//!
//! Returns an `Integer` representing the number of open forms in stand-alone versions of
//! Visual Basic. Returns 0 in all other applications.
//!
//! ## Remarks
//!
//! The `DoEvents` function temporarily yields execution to the operating system, allowing
//! it to process other events such as user input, timers, and system messages. This is
//! essential for keeping an application responsive during long-running operations.
//!
//! **Important Characteristics:**
//!
//! - Yields control to the operating system
//! - Allows message queue processing
//! - Prevents UI from appearing frozen during long operations
//! - Can cause reentrancy issues if not used carefully
//! - Slows down operations slightly due to context switching
//! - Returns number of open forms (VB6 stand-alone only)
//! - In most contexts, return value is ignored
//! - Does not create a new thread or async operation
//! - Processes Windows messages in the queue
//! - Can trigger event handlers and user interactions
//!
//! ## When to Use `DoEvents`
//!
//! - Long-running loops that process data
//! - File operations on large files
//! - Network operations that may take time
//! - Batch processing operations
//! - Any operation that could make UI unresponsive
//!
//! ## When NOT to Use `DoEvents`
//!
//! - In event handlers that could be re-entered
//! - When reentrancy could cause data corruption
//! - In critical sections or transaction code
//! - When better alternatives exist (timers, threading)
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Process large dataset while keeping UI responsive
//! Dim i As Long
//! For i = 1 To 100000
//!     ProcessRecord i
//!     
//!     ' Yield every 100 iterations
//!     If i Mod 100 = 0 Then
//!         DoEvents
//!     End If
//! Next i
//!
//! ' Simple DoEvents call
//! DoEvents
//!
//! ' Check return value (rarely used)
//! Dim formCount As Integer
//! formCount = DoEvents()
//! ```
//!
//! ### Progress Bar Update
//!
//! ```vb
//! Sub ProcessWithProgress()
//!     Dim i As Long
//!     Dim total As Long
//!     
//!     total = 10000
//!     ProgressBar1.Min = 0
//!     ProgressBar1.Max = total
//!     
//!     For i = 1 To total
//!         ProcessItem i
//!         
//!         ' Update progress bar
//!         ProgressBar1.Value = i
//!         lblStatus.Caption = "Processing " & i & " of " & total
//!         
//!         ' Allow UI to refresh
//!         DoEvents
//!     Next i
//!     
//!     MsgBox "Processing complete!"
//! End Sub
//! ```
//!
//! ### File Processing
//!
//! ```vb
//! Sub ProcessLargeFile(filePath As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim lineCount As Long
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     
//!     lineCount = 0
//!     Do Until EOF(fileNum)
//!         Line Input #fileNum, line
//!         ProcessLine line
//!         lineCount = lineCount + 1
//!         
//!         ' Yield every 100 lines
//!         If lineCount Mod 100 = 0 Then
//!             DoEvents
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Cancellable Long Operation
//!
//! ```vb
//! Private cancelOperation As Boolean
//!
//! Sub PerformCancellableOperation()
//!     Dim i As Long
//!     cancelOperation = False
//!     cmdCancel.Enabled = True
//!     
//!     For i = 1 To 100000
//!         If cancelOperation Then
//!             MsgBox "Operation cancelled"
//!             Exit For
//!         End If
//!         
//!         ProcessItem i
//!         
//!         If i Mod 100 = 0 Then
//!             DoEvents  ' Allows cancel button to be clicked
//!         End If
//!     Next i
//!     
//!     cmdCancel.Enabled = False
//! End Sub
//!
//! Private Sub cmdCancel_Click()
//!     cancelOperation = True
//! End Sub
//! ```
//!
//! ### Batch Import with Status
//!
//! ```vb
//! Sub ImportRecords(records As Variant)
//!     Dim i As Long
//!     Dim startTime As Double
//!     
//!     startTime = Timer
//!     
//!     For i = LBound(records) To UBound(records)
//!         ImportRecord records(i)
//!         
//!         ' Update status every 50 records
//!         If i Mod 50 = 0 Then
//!             lblStatus.Caption = "Imported " & i & " records..."
//!             DoEvents
//!         End If
//!     Next i
//!     
//!     lblStatus.Caption = "Import complete: " & UBound(records) - LBound(records) + 1 & _
//!                         " records in " & Format(Timer - startTime, "0.00") & " seconds"
//! End Sub
//! ```
//!
//! ### Prevent UI Freeze During Calculation
//!
//! ```vb
//! Function CalculateComplexResult(data As Variant) As Double
//!     Dim i As Long
//!     Dim result As Double
//!     Dim iterations As Long
//!     
//!     iterations = 0
//!     result = 0
//!     
//!     For i = LBound(data) To UBound(data)
//!         result = result + PerformComplexCalculation(data(i))
//!         iterations = iterations + 1
//!         
//!         ' Yield periodically
//!         If iterations Mod 500 = 0 Then
//!             DoEvents
//!         End If
//!     Next i
//!     
//!     CalculateComplexResult = result
//! End Function
//! ```
//!
//! ### Database Batch Update
//!
//! ```vb
//! Sub UpdateRecordsBatch(rs As ADODB.Recordset)
//!     Dim count As Long
//!     
//!     count = 0
//!     Do Until rs.EOF
//!         rs("Status") = "Processed"
//!         rs("ProcessDate") = Date
//!         rs.Update
//!         
//!         count = count + 1
//!         If count Mod 25 = 0 Then
//!             lblProgress.Caption = count & " records updated"
//!             DoEvents
//!         End If
//!         
//!         rs.MoveNext
//!     Loop
//! End Sub
//! ```
//!
//! ### Search Operation with Live Results
//!
//! ```vb
//! Sub SearchFiles(rootPath As String, searchTerm As String)
//!     Dim fileName As String
//!     Dim matchCount As Long
//!     
//!     matchCount = 0
//!     lstResults.Clear
//!     
//!     fileName = Dir(rootPath & "\*.*")
//!     Do While fileName <> ""
//!         If InStr(1, fileName, searchTerm, vbTextCompare) > 0 Then
//!             lstResults.AddItem fileName
//!             matchCount = matchCount + 1
//!         End If
//!         
//!         fileName = Dir
//!         DoEvents  ' Keep UI responsive, allow viewing results
//!     Loop
//!     
//!     lblStatus.Caption = matchCount & " matches found"
//! End Sub
//! ```
//!
//! ### Report Generation
//!
//! ```vb
//! Sub GenerateReport(data As Collection)
//!     Dim item As Variant
//!     Dim lineNum As Long
//!     
//!     lineNum = 0
//!     
//!     For Each item In data
//!         WriteReportLine item
//!         lineNum = lineNum + 1
//!         
//!         If lineNum Mod 20 = 0 Then
//!             lblProgress.Caption = "Generated " & lineNum & " lines..."
//!             DoEvents
//!         End If
//!     Next item
//! End Sub
//! ```
//!
//! ### Animation or Visual Feedback
//!
//! ```vb
//! Sub ShowProcessingAnimation()
//!     Dim i As Integer
//!     
//!     For i = 1 To 100
//!         ' Update visual indicator
//!         shpIndicator.Left = i * 50
//!         DoEvents
//!         
//!         ' Simulate work
//!         Sleep 10
//!     Next i
//! End Sub
//! ```
//!
//! ### Multi-Step Process
//!
//! ```vb
//! Sub MultiStepProcess()
//!     lblStatus.Caption = "Step 1: Loading data..."
//!     DoEvents
//!     LoadData
//!     
//!     lblStatus.Caption = "Step 2: Processing data..."
//!     DoEvents
//!     ProcessData
//!     
//!     lblStatus.Caption = "Step 3: Saving results..."
//!     DoEvents
//!     SaveResults
//!     
//!     lblStatus.Caption = "Complete!"
//!     DoEvents
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Prevent Reentrancy
//!
//! ```vb
//! Private isProcessing As Boolean
//!
//! Sub SafeProcessWithDoEvents()
//!     ' Prevent re-entry
//!     If isProcessing Then
//!         MsgBox "Already processing"
//!         Exit Sub
//!     End If
//!     
//!     isProcessing = True
//!     
//!     Dim i As Long
//!     For i = 1 To 10000
//!         ProcessItem i
//!         
//!         If i Mod 100 = 0 Then
//!             DoEvents
//!         End If
//!     Next i
//!     
//!     isProcessing = False
//! End Sub
//! ```
//!
//! ### Throttled `DoEvents`
//!
//! ```vb
//! Sub ProcessWithThrottledDoEvents()
//!     Dim i As Long
//!     Dim lastDoEvents As Double
//!     Dim doEventsInterval As Double
//!     
//!     doEventsInterval = 0.1  ' 100ms
//!     lastDoEvents = Timer
//!     
//!     For i = 1 To 100000
//!         ProcessItem i
//!         
//!         ' DoEvents based on time, not iteration count
//!         If Timer - lastDoEvents > doEventsInterval Then
//!             DoEvents
//!             lastDoEvents = Timer
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ### Disable Controls During Processing
//!
//! ```vb
//! Sub ProcessWithDisabledControls()
//!     ' Disable controls to prevent reentrancy
//!     DisableControls
//!     
//!     Dim i As Long
//!     For i = 1 To 10000
//!         ProcessItem i
//!         
//!         If i Mod 100 = 0 Then
//!             UpdateProgress i
//!             DoEvents  ' Safe because controls are disabled
//!         End If
//!     Next i
//!     
//!     EnableControls
//! End Sub
//!
//! Sub DisableControls()
//!     Dim ctrl As Control
//!     For Each ctrl In Me.Controls
//!         If TypeOf ctrl Is CommandButton Then
//!             ctrl.Enabled = False
//!         End If
//!     Next ctrl
//! End Sub
//!
//! Sub EnableControls()
//!     Dim ctrl As Control
//!     For Each ctrl In Me.Controls
//!         If TypeOf ctrl Is CommandButton Then
//!             ctrl.Enabled = True
//!         End If
//!     Next ctrl
//! End Sub
//! ```
//!
//! ### Background Processing Simulation
//!
//! ```vb
//! ' Simulates background processing using DoEvents
//! Private processingComplete As Boolean
//!
//! Sub StartBackgroundTask()
//!     processingComplete = False
//!     
//!     ' Start the "background" task
//!     ProcessInBackground
//!     
//!     ' Show modal dialog that waits
//!     Do Until processingComplete
//!         DoEvents
//!         Sleep 10  ' Small delay to reduce CPU usage
//!     Loop
//!     
//!     MsgBox "Background task complete"
//! End Sub
//!
//! Sub ProcessInBackground()
//!     Dim i As Long
//!     For i = 1 To 10000
//!         ProcessItem i
//!         
//!         If i Mod 100 = 0 Then
//!             DoEvents
//!         End If
//!     Next i
//!     
//!     processingComplete = True
//! End Sub
//! ```
//!
//! ### Smart `DoEvents` with CPU Management
//!
//! ```vb
//! Sub ProcessWithCPUManagement()
//!     Dim i As Long
//!     Dim processingTime As Double
//!     Dim doEventsTime As Double
//!     
//!     For i = 1 To 100000
//!         processingTime = Timer
//!         ProcessItem i
//!         processingTime = Timer - processingTime
//!         
//!         ' DoEvents if processing takes significant time
//!         If processingTime > 0.05 Then  ' More than 50ms
//!             doEventsTime = Timer
//!             DoEvents
//!             doEventsTime = Timer - doEventsTime
//!             
//!             ' Adjust strategy if DoEvents takes too long
//!             If doEventsTime > processingTime * 0.1 Then
//!                 ' DoEvents overhead is too high, reduce frequency
//!                 ' (implementation-specific logic here)
//!             End If
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ### Export with User Interaction
//!
//! ```vb
//! Sub ExportDataWithOptions()
//!     Dim i As Long
//!     Dim exportCount As Long
//!     
//!     exportCount = 0
//!     
//!     For i = 1 To RecordCount
//!         If chkIncludeDeleted.Value = vbChecked Or Not IsDeleted(i) Then
//!             ExportRecord i
//!             exportCount = exportCount + 1
//!         End If
//!         
//!         ' Update UI and allow user to change options
//!         If i Mod 50 = 0 Then
//!             lblProgress.Caption = exportCount & " records exported"
//!             DoEvents
//!             ' User can check/uncheck options, affecting subsequent exports
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Sub ProcessWithErrorHandling()
//!     On Error GoTo ErrorHandler
//!     
//!     Dim i As Long
//!     For i = 1 To 10000
//!         ProcessItem i
//!         
//!         If i Mod 100 = 0 Then
//!             DoEvents
//!         End If
//!     Next i
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     ' DoEvents can trigger errors if event handlers fail
//!     MsgBox "Error during processing: " & Err.Description
//!     Resume Next
//! End Sub
//! ```
//!
//! ### Common Errors
//!
//! - **Error 11** (Division by zero): Can occur if `DoEvents` allows user to clear data
//! - **Error 91** (Object variable not set): If `DoEvents` allows object to be destroyed
//! - **Reentrancy errors**: If `DoEvents` allows same code to be called recursively
//!
//! ## Performance Considerations
//!
//! - `DoEvents` has overhead (context switching, message processing)
//! - Call too frequently: significant performance impact
//! - Call too infrequently: UI appears frozen
//! - Typical guideline: every 50-100 iterations or every 100ms
//! - For very fast loops, use time-based checking instead of iteration-based
//! - Consider alternatives for truly asynchronous operations
//! - `Sleep()` between `DoEvents` can reduce CPU usage in wait loops
//!
//! ## Best Practices
//!
//! ### Call Periodically, Not Every Iteration
//!
//! ```vb
//! ' Good - DoEvents every 100 iterations
//! For i = 1 To 100000
//!     ProcessItem i
//!     If i Mod 100 = 0 Then DoEvents
//! Next i
//!
//! ' Bad - DoEvents every iteration (slow)
//! For i = 1 To 100000
//!     ProcessItem i
//!     DoEvents  ' Too frequent!
//! Next i
//! ```
//!
//! ### Protect Against Reentrancy
//!
//! ```vb
//! ' Good - Use flag to prevent reentrancy
//! Private isProcessing As Boolean
//!
//! Sub Process()
//!     If isProcessing Then Exit Sub
//!     isProcessing = True
//!     ' ... processing with DoEvents ...
//!     isProcessing = False
//! End Sub
//!
//! ' Bad - No protection
//! Sub Process()
//!     ' ... processing with DoEvents ...
//!     ' Can be called again through DoEvents
//! End Sub
//! ```
//!
//! ### Disable User Input When Needed
//!
//! ```vb
//! ' Good - Disable controls that could cause problems
//! cmdProcess.Enabled = False
//! For i = 1 To 10000
//!     ProcessItem i
//!     If i Mod 100 = 0 Then DoEvents
//! Next i
//! cmdProcess.Enabled = True
//! ```
//!
//! ### Consider Alternatives
//!
//! ```vb
//! ' For very long operations, consider:
//! ' 1. Timer control for asynchronous processing
//! ' 2. Threading (in modern applications)
//! ' 3. Breaking into smaller chunks with callbacks
//! ' 4. Progress forms with asynchronous updates
//! ```
//!
//! ## Comparison with Other Approaches
//!
//! ### `DoEvents` vs Timer Control
//!
//! ```vb
//! ' DoEvents - Synchronous, blocks until complete
//! For i = 1 To 10000
//!     ProcessItem i
//!     If i Mod 100 = 0 Then DoEvents
//! Next i
//!
//! ' Timer - Asynchronous, processes in chunks
//! Private currentIndex As Long
//!
//! Private Sub Timer1_Timer()
//!     Dim i As Long
//!     For i = currentIndex To currentIndex + 99
//!         If i > 10000 Then
//!             Timer1.Enabled = False
//!             Exit Sub
//!         End If
//!         ProcessItem i
//!     Next i
//!     currentIndex = i
//! End Sub
//! ```
//!
//! ### `DoEvents` vs `Application.Wait` (Excel VBA)
//!
//! ```vb
//! ' DoEvents - Yields immediately
//! DoEvents
//!
//! ' Application.Wait - Yields for specific duration
//! Application.Wait Now + TimeValue("00:00:01")
//! ```
//!
//! ## Limitations
//!
//! - Does not create true multithreading
//! - Can cause reentrancy issues
//! - Performance overhead from context switching
//! - Return value rarely useful in modern applications
//! - No control over which events are processed
//! - Can make debugging more difficult
//! - Not available in all VBA hosts (e.g., some Office apps)
//! - Cannot cancel or prioritize specific events
//!
//! ## Related Functions
//!
//! - `Sleep`: Pauses execution for specified milliseconds (Windows API)
//! - `Timer`: Returns seconds since midnight (for timing operations)
//! - `Now`: Returns current date and time
//! - `Wait`: Application-specific wait method (Excel VBA)
//! - **Timer Control**: Asynchronous event-based processing
//! - **Threading APIs**: True multithreading (advanced)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn doevents_basic() {
        let source = r#"
DoEvents
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_parentheses() {
        let source = r#"
DoEvents()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_in_loop() {
        let source = r#"
For i = 1 To 10000
    ProcessItem i
    If i Mod 100 = 0 Then DoEvents
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_assignment() {
        let source = r#"
formCount = DoEvents()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_in_do_loop() {
        let source = r#"
Do Until EOF(1)
    ProcessLine line
    DoEvents
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_cancellable_operation() {
        let source = r#"
For i = 1 To 100000
    If cancelOperation Then Exit For
    ProcessItem i
    If i Mod 100 = 0 Then DoEvents
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_status_update() {
        let source = r#"
lblStatus.Caption = "Processing..."
DoEvents
ProcessData
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_progress_bar() {
        let source = r#"
ProgressBar1.Value = i
DoEvents
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_while_loop() {
        let source = r#"
Do While fileName <> ""
    ProcessFile fileName
    fileName = Dir
    DoEvents
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_batch_update() {
        let source = r#"
Do Until rs.EOF
    rs.Update
    If count Mod 25 = 0 Then DoEvents
    rs.MoveNext
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_multi_step() {
        let source = r#"
lblStatus.Caption = "Step 1"
DoEvents
LoadData
lblStatus.Caption = "Step 2"
DoEvents
ProcessData
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_in_function() {
        let source = r#"
Function ProcessData() As Boolean
    Dim i As Long
    For i = 1 To 1000
        ProcessItem i
        DoEvents
    Next i
    ProcessData = True
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_error_handling() {
        let source = r#"
On Error Resume Next
For i = 1 To 10000
    ProcessItem i
    DoEvents
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_reentrancy_guard() {
        let source = r#"
If isProcessing Then Exit Sub
isProcessing = True
For i = 1 To 1000
    DoEvents
Next i
isProcessing = False
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_file_processing() {
        let source = r#"
Do Until EOF(fileNum)
    Line Input #fileNum, line
    ProcessLine line
    lineCount = lineCount + 1
    If lineCount Mod 100 = 0 Then DoEvents
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_conditional() {
        let source = r#"
If Timer - lastUpdate > 0.1 Then
    DoEvents
    lastUpdate = Timer
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_select_case() {
        let source = r#"
Select Case step
    Case 1
        x = DoEvents()
        ProcessStep1
    Case 2
        y = DoEvents()
        ProcessStep2
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_nested_loop() {
        let source = r#"
For i = 1 To 100
    For j = 1 To 100
        Process i, j
    Next j
    DoEvents
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_sleep() {
        let source = r#"
Do Until processingComplete
    DoEvents
    Sleep 10
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_search_operation() {
        let source = r#"
fileName = Dir("*.*")
Do While fileName <> ""
    If InStr(fileName, searchTerm) > 0 Then
        lstResults.AddItem fileName
    End If
    DoEvents
    fileName = Dir
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_animation() {
        let source = r#"
For i = 1 To 100
    shpIndicator.Left = i * 50
    DoEvents
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_inline_if() {
        let source = r#"
If i Mod 100 = 0 Then DoEvents Else ProcessFast
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_record_processing() {
        let source = r#"
For Each item In collection
    ProcessItem item
    DoEvents
Next item
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_export() {
        let source = r#"
For i = 1 To recordCount
    ExportRecord i
    If i Mod 50 = 0 Then
        lblProgress.Caption = i & " exported"
        DoEvents
    End If
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn doevents_with_call() {
        let source = r#"
Call DoEvents
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DoEvents"));
        assert!(debug.contains("Identifier"));
    }
}
