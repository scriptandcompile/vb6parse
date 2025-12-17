//! # `Error` Function
//!
//! Returns the error message that corresponds to a given error number.
//!
//! ## Syntax
//!
//! ```vb
//! Error[(errornumber)]
//! ```
//!
//! ## Parameters
//!
//! - **errornumber**: Optional. A Long or any valid numeric expression that represents
//!   an error number. If omitted, the error message for the most recent run-time error
//!   (the current value of `Err.Number`) is returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the error message associated with the specified error number.
//! If the error number is not recognized, `Error` returns "Application-defined or object-defined error".
//!
//! ## Remarks
//!
//! The `Error` function is used to retrieve the text description of VB6 run-time errors.
//! It's useful for error handling, logging, and displaying user-friendly error messages.
//!
//! **Important Characteristics:**
//!
//! - Returns error message as `String`
//! - Without argument, returns message for current error (`Err.Number`)
//! - With argument, returns message for specified error number
//! - VB6 error numbers range from 0 to 65535
//! - User-defined errors typically use 512-65535 range
//! - System errors use 0-511 range
//! - Unrecognized errors return generic message
//! - Does not clear or raise errors
//! - Can be used without On Error statement
//! - `Err.Description` also provides error messages
//!
//! ## Common VB6 Error Numbers
//!
//! - **3**: Return without `GoSub`
//! - **5**: Invalid procedure call or argument
//! - **6**: Overflow
//! - **7**: Out of memory
//! - **9**: Subscript out of range
//! - **10**: Array is fixed or temporarily locked
//! - **11**: Division by zero
//! - **13**: Type mismatch
//! - **28**: Out of stack space
//! - **35**: Sub or Function not defined
//! - **48**: Error in loading DLL
//! - **49**: Bad DLL calling convention
//! - **51**: Internal error
//! - **52**: Bad file name or number
//! - **53**: File not found
//! - **54**: Bad file mode
//! - **55**: File already open
//! - **57**: Device I/O error
//! - **58**: File already exists
//! - **61**: Disk full
//! - **62**: Input past end of file
//! - **63**: Bad record number
//! - **67**: Too many files
//! - **68**: Device unavailable
//! - **70**: Permission denied
//! - **71**: Disk not ready
//! - **74**: Can't rename with different drive
//! - **75**: Path/File access error
//! - **76**: Path not found
//! - **91**: `Object` variable or `With` block variable not set
//! - **92**: For loop not initialized
//! - **93**: Invalid pattern string
//! - **94**: Invalid use of Null
//! - **298**: System DLL could not be loaded
//! - **321**: Invalid file format
//! - **322**: Can't create necessary temporary file
//! - **325**: Invalid format in resource file
//! - **380**: Invalid property value
//! - **424**: `Object` required
//! - **429**: `ActiveX` component can't create `Object`
//! - **430**: Class does not support Automation
//! - **432**: File name or class name not found during Automation operation
//! - **438**: `Object` doesn't support this property or method
//! - **440**: Automation error
//! - **445**: `Object` doesn't support this action
//! - **446**: `Object` doesn't support named arguments
//! - **447**: `Object` doesn't support current locale setting
//! - **448**: Named argument not found
//! - **449**: Argument not optional
//! - **450**: Wrong number of arguments or invalid property assignment
//! - **451**: `Object` not a collection
//! - **452**: Invalid ordinal
//! - **453**: Specified DLL function not found
//! - **454**: Code resource not found
//! - **455**: Code resource lock error
//! - **457**: This key is already associated with an element of this collection
//! - **458**: Variable uses an Automation type not supported in Visual Basic
//! - **459**: This component doesn't support events
//! - **460**: Invalid Clipboard format
//! - **461**: Specified format doesn't match format of data
//! - **480**: Can't create `AutoRedraw` image
//! - **481**: Invalid picture
//! - **482**: Printer error
//! - **735**: Can't save file to TEMP
//! - **744**: Search text not found
//! - **746**: Replacements too long
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get error message for specific error number
//! Dim msg As String
//! msg = Error(53)
//! Debug.Print msg  ' Prints: "File not found"
//!
//! msg = Error(11)
//! Debug.Print msg  ' Prints: "Division by zero"
//! ```
//!
//! ### Get Current Error Message
//!
//! ```vb
//! Sub TestErrorHandling()
//!     On Error Resume Next
//!     
//!     ' Cause an error
//!     Dim x As Integer
//!     x = 1 / 0
//!     
//!     ' Check if error occurred
//!     If Err.Number <> 0 Then
//!         ' Get error message without parameter (uses current error)
//!         Debug.Print "Error: " & Error
//!         Debug.Print "Error number: " & Err.Number
//!         Err.Clear
//!     End If
//! End Sub
//! ```
//!
//! ### Display Error in Message Box
//!
//! ```vb
//! Sub OpenFileWithErrorHandling(filePath As String)
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileNum As Integer
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     ' Process file...
//!     Close #fileNum
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error " & Err.Number & ": " & Error(Err.Number), _
//!            vbExclamation, "File Error"
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Error Lookup Table
//!
//! ```vb
//! Function GetErrorMessage(errorNumber As Long) As String
//!     ' Get standard VB6 error message
//!     GetErrorMessage = Error(errorNumber)
//!     
//!     ' Override with custom messages if desired
//!     Select Case errorNumber
//!         Case 53
//!             GetErrorMessage = "The specified file could not be found. Please check the path."
//!         Case 61
//!             GetErrorMessage = "The disk is full. Please free up space and try again."
//!         Case 91
//!             GetErrorMessage = "Object variable not initialized. Please contact support."
//!     End Select
//! End Function
//! ```
//!
//! ### Error Logging
//!
//! ```vb
//! Sub LogError(procedureName As String)
//!     Dim fileNum As Integer
//!     Dim logPath As String
//!     
//!     logPath = App.Path & "\error.log"
//!     fileNum = FreeFile
//!     
//!     Open logPath For Append As #fileNum
//!     Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
//!                     procedureName & " | " & _
//!                     "Error " & Err.Number & ": " & Error(Err.Number)
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Custom Error Messages
//!
//! ```vb
//! Function GetFriendlyErrorMessage(errorNumber As Long) As String
//!     Dim standardMsg As String
//!     standardMsg = Error(errorNumber)
//!     
//!     ' Provide user-friendly alternatives
//!     Select Case errorNumber
//!         Case 53  ' File not found
//!             GetFriendlyErrorMessage = "We couldn't find that file. " & _
//!                 "It may have been moved or deleted."
//!         
//!         Case 70  ' Permission denied
//!             GetFriendlyErrorMessage = "You don't have permission to access this file. " & _
//!                 "Please contact your administrator."
//!         
//!         Case 429  ' Can't create object
//!             GetFriendlyErrorMessage = "A required component is not installed. " & _
//!                 "Please reinstall the application."
//!         
//!         Case Else
//!             GetFriendlyErrorMessage = standardMsg
//!     End Select
//! End Function
//! ```
//!
//! ### Error Report Dialog
//!
//! ```vb
//! Sub ShowErrorReport()
//!     Dim msg As String
//!     
//!     msg = "An error has occurred:" & vbCrLf & vbCrLf
//!     msg = msg & "Error Number: " & Err.Number & vbCrLf
//!     msg = msg & "Description: " & Error(Err.Number) & vbCrLf
//!     msg = msg & "Source: " & Err.Source & vbCrLf
//!     
//!     If Err.HelpFile <> "" Then
//!         msg = msg & "Help File: " & Err.HelpFile & vbCrLf
//!         msg = msg & "Help Context: " & Err.HelpContext & vbCrLf
//!     End If
//!     
//!     MsgBox msg, vbCritical, "Application Error"
//! End Sub
//! ```
//!
//! ### Validate Error Numbers
//!
//! ```vb
//! Function IsValidErrorNumber(errNum As Long) As Boolean
//!     Dim msg As String
//!     
//!     ' Get error message
//!     msg = Error(errNum)
//!     
//!     ' If it's not a recognized error, VB6 returns a generic message
//!     If InStr(msg, "Application-defined") > 0 Or _
//!        InStr(msg, "object-defined") > 0 Then
//!         IsValidErrorNumber = False
//!     Else
//!         IsValidErrorNumber = True
//!     End If
//! End Function
//! ```
//!
//! ### List Common Errors
//!
//! ```vb
//! Sub ListCommonErrors()
//!     Dim errorNumbers() As Long
//!     Dim i As Integer
//!     
//!     errorNumbers = Array(5, 6, 7, 9, 11, 13, 52, 53, 54, 61, 62, 70, 91, 429)
//!     
//!     Debug.Print "Common VB6 Errors:"
//!     Debug.Print String(50, "-")
//!     
//!     For i = LBound(errorNumbers) To UBound(errorNumbers)
//!         Debug.Print errorNumbers(i) & ": " & Error(errorNumbers(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Enhanced Error Handler
//!
//! ```vb
//! Function HandleError(moduleName As String, procedureName As String) As VbMsgBoxResult
//!     Dim msg As String
//!     Dim errorMsg As String
//!     
//!     errorMsg = Error(Err.Number)
//!     
//!     msg = "An error occurred in " & moduleName & "." & procedureName & vbCrLf & vbCrLf
//!     msg = msg & "Error " & Err.Number & ": " & errorMsg & vbCrLf & vbCrLf
//!     msg = msg & "Would you like to continue?"
//!     
//!     HandleError = MsgBox(msg, vbYesNo + vbExclamation, "Error")
//!     
//!     ' Log the error
//!     LogError moduleName & "." & procedureName
//!     
//!     ' Clear the error
//!     Err.Clear
//! End Function
//! ```
//!
//! ### Error Dictionary
//!
//! ```vb
//! Function BuildErrorDictionary() As Collection
//!     Dim dict As New Collection
//!     Dim i As Long
//!     Dim msg As String
//!     
//!     ' Build dictionary of all valid error messages
//!     For i = 3 To 1000
//!         msg = Error(i)
//!         
//!         ' Only add if it's a recognized error
//!         If InStr(msg, "Application-defined") = 0 Then
//!             On Error Resume Next
//!             dict.Add msg, CStr(i)
//!             On Error GoTo 0
//!         End If
//!     Next i
//!     
//!     Set BuildErrorDictionary = dict
//! End Function
//! ```
//!
//! ### Multilingual Error Messages
//!
//! ```vb
//! Function GetLocalizedError(errorNumber As Long, language As String) As String
//!     Dim standardMsg As String
//!     standardMsg = Error(errorNumber)
//!     
//!     ' Provide translations for common errors
//!     If language = "ES" Then  ' Spanish
//!         Select Case errorNumber
//!             Case 53: GetLocalizedError = "Archivo no encontrado"
//!             Case 61: GetLocalizedError = "Disco lleno"
//!             Case 70: GetLocalizedError = "Permiso denegado"
//!             Case Else: GetLocalizedError = standardMsg
//!         End Select
//!     ElseIf language = "FR" Then  ' French
//!         Select Case errorNumber
//!             Case 53: GetLocalizedError = "Fichier non trouvé"
//!             Case 61: GetLocalizedError = "Disque plein"
//!             Case 70: GetLocalizedError = "Permission refusée"
//!             Case Else: GetLocalizedError = standardMsg
//!         End Select
//!     Else
//!         GetLocalizedError = standardMsg
//!     End If
//! End Function
//! ```
//!
//! ### Error Testing Helper
//!
//! ```vb
//! Sub TestErrorMessages()
//!     Dim testErrors() As Long
//!     Dim i As Integer
//!     
//!     ' Test specific error numbers
//!     testErrors = Array(5, 6, 7, 9, 11, 13, 52, 53, 54, 55, 57, 58, 61, 62, _
//!                        67, 68, 70, 71, 74, 75, 76, 91, 92, 93, 94)
//!     
//!     For i = LBound(testErrors) To UBound(testErrors)
//!         Debug.Print "Error " & testErrors(i) & ": " & Error(testErrors(i))
//!     Next i
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Error Mapper with Suggestions
//!
//! ```vb
//! Type ErrorInfo
//!     Number As Long
//!     Message As String
//!     Suggestion As String
//! End Type
//!
//! Function GetErrorInfo(errorNumber As Long) As ErrorInfo
//!     Dim info As ErrorInfo
//!     
//!     info.Number = errorNumber
//!     info.Message = Error(errorNumber)
//!     
//!     ' Add helpful suggestions
//!     Select Case errorNumber
//!         Case 53
//!             info.Suggestion = "Check the file path and ensure the file exists."
//!         Case 61
//!             info.Suggestion = "Free up disk space or choose a different location."
//!         Case 70
//!             info.Suggestion = "Run the application as administrator or check file permissions."
//!         Case 91
//!             info.Suggestion = "Ensure the object is initialized with New or Set."
//!         Case 429
//!             info.Suggestion = "Verify that all required DLLs and components are registered."
//!         Case Else
//!             info.Suggestion = "Please contact technical support if the problem persists."
//!     End Select
//!     
//!     GetErrorInfo = info
//! End Function
//! ```
//!
//! ### Comprehensive Error Logger
//!
//! ```vb
//! Sub LogDetailedError(moduleName As String, procedureName As String, _
//!                      Optional additionalInfo As String = "")
//!     Dim fileNum As Integer
//!     Dim logPath As String
//!     Dim logEntry As String
//!     
//!     logPath = App.Path & "\logs\error_" & Format(Now, "yyyymmdd") & ".log"
//!     
//!     ' Build detailed log entry
//!     logEntry = String(80, "=") & vbCrLf
//!     logEntry = logEntry & "Timestamp: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
//!     logEntry = logEntry & "Module: " & moduleName & vbCrLf
//!     logEntry = logEntry & "Procedure: " & procedureName & vbCrLf
//!     logEntry = logEntry & "Error Number: " & Err.Number & vbCrLf
//!     logEntry = logEntry & "Error Message: " & Error(Err.Number) & vbCrLf
//!     logEntry = logEntry & "Error Source: " & Err.Source & vbCrLf
//!     
//!     If additionalInfo <> "" Then
//!         logEntry = logEntry & "Additional Info: " & additionalInfo & vbCrLf
//!     End If
//!     
//!     logEntry = logEntry & String(80, "=") & vbCrLf & vbCrLf
//!     
//!     ' Write to log file
//!     fileNum = FreeFile
//!     Open logPath For Append As #fileNum
//!     Print #fileNum, logEntry
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Error Recovery System
//!
//! ```vb
//! Function AttemptRecovery(errorNumber As Long) As Boolean
//!     Dim errorMsg As String
//!     
//!     errorMsg = Error(errorNumber)
//!     AttemptRecovery = False
//!     
//!     Select Case errorNumber
//!         Case 53  ' File not found
//!             ' Try to create the file or directory
//!             On Error Resume Next
//!             MkDir GetParentDirectory(expectedFilePath)
//!             CreateDefaultFile expectedFilePath
//!             If Err.Number = 0 Then AttemptRecovery = True
//!             On Error GoTo 0
//!         
//!         Case 61  ' Disk full
//!             ' Try to clean temp files
//!             On Error Resume Next
//!             CleanTempFiles
//!             If GetFreeDiskSpace() > 1048576 Then AttemptRecovery = True
//!             On Error GoTo 0
//!         
//!         Case 70  ' Permission denied
//!             ' Prompt user to run as administrator
//!             MsgBox "This operation requires administrator privileges. " & _
//!                    "Please restart the application as administrator.", vbExclamation
//!             AttemptRecovery = False
//!         
//!         Case 91  ' Object not set
//!             ' Try to reinitialize object
//!             On Error Resume Next
//!             InitializeObjects
//!             If Err.Number = 0 Then AttemptRecovery = True
//!             On Error GoTo 0
//!     End Select
//! End Function
//! ```
//!
//! ### Error Statistics Tracker
//!
//! ```vb
//! Type ErrorStat
//!     ErrorNumber As Long
//!     ErrorMessage As String
//!     OccurrenceCount As Long
//!     LastOccurrence As Date
//! End Type
//!
//! Private errorStats() As ErrorStat
//! Private statCount As Long
//!
//! Sub TrackError(errorNumber As Long)
//!     Dim i As Long
//!     Dim found As Boolean
//!     
//!     ' Find existing error in stats
//!     found = False
//!     For i = 0 To statCount - 1
//!         If errorStats(i).ErrorNumber = errorNumber Then
//!             errorStats(i).OccurrenceCount = errorStats(i).OccurrenceCount + 1
//!             errorStats(i).LastOccurrence = Now
//!             found = True
//!             Exit For
//!         End If
//!     Next i
//!     
//!     ' Add new error to stats
//!     If Not found Then
//!         ReDim Preserve errorStats(0 To statCount)
//!         errorStats(statCount).ErrorNumber = errorNumber
//!         errorStats(statCount).ErrorMessage = Error(errorNumber)
//!         errorStats(statCount).OccurrenceCount = 1
//!         errorStats(statCount).LastOccurrence = Now
//!         statCount = statCount + 1
//!     End If
//! End Sub
//!
//! Function GetErrorStatistics() As String
//!     Dim i As Long
//!     Dim report As String
//!     
//!     report = "Error Statistics Report" & vbCrLf
//!     report = report & String(80, "-") & vbCrLf & vbCrLf
//!     
//!     For i = 0 To statCount - 1
//!         With errorStats(i)
//!             report = report & "Error " & .ErrorNumber & ": " & .ErrorMessage & vbCrLf
//!             report = report & "  Occurrences: " & .OccurrenceCount & vbCrLf
//!             report = report & "  Last Seen: " & Format(.LastOccurrence, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
//!         End With
//!     Next i
//!     
//!     GetErrorStatistics = report
//! End Function
//! ```
//!
//! ### Email Error Notification
//!
//! ```vb
//! Sub SendErrorNotification(errorNumber As Long, context As String)
//!     Dim emailBody As String
//!     Dim errorMsg As String
//!     
//!     errorMsg = Error(errorNumber)
//!     
//!     emailBody = "An error occurred in the application:" & vbCrLf & vbCrLf
//!     emailBody = emailBody & "Error Number: " & errorNumber & vbCrLf
//!     emailBody = emailBody & "Error Message: " & errorMsg & vbCrLf
//!     emailBody = emailBody & "Context: " & context & vbCrLf
//!     emailBody = emailBody & "User: " & Environ("USERNAME") & vbCrLf
//!     emailBody = emailBody & "Computer: " & Environ("COMPUTERNAME") & vbCrLf
//!     emailBody = emailBody & "Timestamp: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
//!     
//!     ' Send email (pseudo-code)
//!     SendEmail "admin@company.com", "Application Error", emailBody
//! End Sub
//! ```
//!
//! ### Error-Based Retry Logic
//!
//! ```vb
//! Function ExecuteWithRetry(operation As String, maxRetries As Integer) As Boolean
//!     Dim retryCount As Integer
//!     Dim errorMsg As String
//!     
//!     retryCount = 0
//!     
//!     Do
//!         On Error Resume Next
//!         
//!         ' Attempt operation
//!         ExecuteOperation operation
//!         
//!         If Err.Number = 0 Then
//!             ExecuteWithRetry = True
//!             Exit Function
//!         End If
//!         
//!         ' Get error message
//!         errorMsg = Error(Err.Number)
//!         
//!         ' Check if error is retryable
//!         Select Case Err.Number
//!             Case 57, 68  ' Device I/O error, Device unavailable
//!                 retryCount = retryCount + 1
//!                 If retryCount < maxRetries Then
//!                     Debug.Print "Retry " & retryCount & " after error: " & errorMsg
//!                     Sleep 1000  ' Wait before retry
//!                 End If
//!             
//!             Case Else  ' Non-retryable error
//!                 Debug.Print "Non-retryable error: " & errorMsg
//!                 Exit Do
//!         End Select
//!         
//!         On Error GoTo 0
//!     Loop While retryCount < maxRetries
//!     
//!     ExecuteWithRetry = False
//! End Function
//! ```
//!
//! ## Error Handling Best Practices
//!
//! ```vb
//! ' Good - Use Error function for logging and display
//! Sub ProcessFile(filePath As String)
//!     On Error GoTo ErrorHandler
//!     
//!     ' Processing code...
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     Dim errorMsg As String
//!     errorMsg = Error(Err.Number)
//!     
//!     LogError "ProcessFile", errorMsg
//!     MsgBox "Failed to process file: " & errorMsg, vbExclamation
//!     Err.Clear
//! End Sub
//!
//! ' Good - Compare with Err.Description
//! Sub CompareErrorSources()
//!     On Error Resume Next
//!     Dim x As Integer
//!     x = 1 / 0
//!     
//!     Debug.Print "Error function: " & Error(Err.Number)
//!     Debug.Print "Err.Description: " & Err.Description
//!     ' Both typically return the same message
//! End Sub
//! ```
//!
//! ## Performance Considerations
//!
//! - `Error` function is very fast (simple lookup)
//! - No performance difference between `Error()` and `Error(n)`
//! - Message strings are pre-defined in VB6 runtime
//! - Consider caching messages if calling repeatedly
//! - Minimal overhead for error message retrieval
//!
//! ## Comparison with Other Error Functions
//!
//! ### `Error` vs `Err.Description`
//!
//! ```vb
//! ' Error() - Returns message for specified or current error
//! msg = Error(53)          ' "File not found"
//! msg = Error()            ' Current error message
//!
//! ' Err.Description - Always current error message
//! msg = Err.Description    ' Current error message only
//! ```
//!
//! ### `Error` vs `Err.Raise`
//!
//! ```vb
//! ' Error() - Retrieves error message (does not raise)
//! msg = Error(5)           ' Just gets the message
//!
//! ' Err.Raise - Raises an error
//! Err.Raise 5              ' Triggers error 5
//! ```
//!
//! ## Limitations
//!
//! - Returns only standard VB6 error messages
//! - Cannot customize built-in messages
//! - Unrecognized error numbers return generic message
//! - Does not provide error context or call stack
//! - Limited to VB6 error number range
//! - No support for system error codes directly
//!
//! ## Related Functions
//!
//! - `Err.Number`: Returns the current error number
//! - `Err.Description`: Returns the current error description
//! - `Err.Raise`: Raises a run-time error
//! - `Err.Clear`: Clears current error information
//! - `CVErr`: Creates an error value
//! - `IsError`: Tests if a variant contains an error value

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn error_basic() {
        let source = r#"
msg = Error(53)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_no_argument() {
        let source = r#"
msg = Error()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_in_msgbox() {
        let source = r#"
MsgBox "Error: " & Error(Err.Number)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_debug_print() {
        let source = r#"
Debug.Print Error(11)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_in_function() {
        let source = r#"
Function GetErrorMessage(errNum As Long) As String
    GetErrorMessage = Error(errNum)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_variable() {
        let source = r#"
errorMsg = Error(errorNumber)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_concatenation() {
        let source = r#"
msg = "Error " & Err.Number & ": " & Error(Err.Number)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_select_case() {
        let source = r#"
Select Case Error(errNum)
    Case "File not found"
        HandleFileNotFound
    Case Else
        HandleOtherError
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_in_if() {
        let source = r#"
If InStr(Error(errNum), "Application-defined") > 0 Then
    isCustomError = True
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_file_logging() {
        let source = r#"
Print #fileNum, "Error " & Err.Number & ": " & Error(Err.Number)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_format() {
        let source = r#"
logEntry = Format(Now, "yyyy-mm-dd") & " | " & Error(Err.Number)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_array_lookup() {
        let source = r#"
For i = LBound(errorNumbers) To UBound(errorNumbers)
    Debug.Print errorNumbers(i) & ": " & Error(errorNumbers(i))
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_collection_add() {
        let source = r#"
dict.Add Error(i), CStr(i)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_comparison() {
        let source = r#"
msg = Error(53)
If msg = "File not found" Then
    Debug.Print "Match"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_udt_field() {
        let source = r#"
info.Message = Error(errorNumber)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_multiline_expression() {
        let source = r#"
msg = "An error occurred:" & vbCrLf & _
      "Error Number: " & Err.Number & vbCrLf & _
      "Description: " & Error(Err.Number)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_left() {
        let source = r#"
shortMsg = Left(Error(errNum), 50)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_nested_call() {
        let source = r#"
logEntry = Replace(Error(Err.Number), vbCrLf, " ")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_len() {
        let source = r#"
If Len(Error(errNum)) > 0 Then
    ProcessError
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_type_assignment() {
        let source = r#"
Dim errorInfo As ErrorInfo
errorInfo.Message = Error(errorNumber)
errorInfo.Number = errorNumber
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_loop_variable() {
        let source = r#"
For i = 3 To 1000
    msg = Error(i)
    If InStr(msg, "Application-defined") = 0 Then
        Debug.Print i, msg
    End If
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_trim() {
        let source = r#"
cleanMsg = Trim(Error(errNum))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_case_statement() {
        let source = r#"
Select Case errNum
    Case 53
        msg = Error(errNum)
    Case 61
        msg = Error(61)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_with_ucase() {
        let source = r#"
upperMsg = UCase(Error(errNum))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_immediate_print() {
        let source = r#"
Sub ShowError()
    Debug.Print Error
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn error_return_value() {
        let source = r#"
Function GetMsg() As String
    GetMsg = Error()
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Error"));
        assert!(debug.contains("Identifier"));
    }
}
