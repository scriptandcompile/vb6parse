//! # Now Function
//!
//! Returns a Variant (Date) specifying the current date and time according to the setting of the computer's system date and time.
//!
//! ## Syntax
//!
//! ```vb
//! Now
//! ```
//!
//! ## Parameters
//!
//! None. Now takes no parameters.
//!
//! ## Return Value
//!
//! Returns a **Variant (Date)** containing the current system date and time.
//!
//! The returned value includes both the date portion (number of days since December 30, 1899) and the time portion (fractional part of a 24-hour day).
//!
//! ## Remarks
//!
//! The Now function is one of the most frequently used VB6 date/time functions for getting the current moment.
//! It combines both date and time information in a single value.
//!
//! ### Key Characteristics:
//! - Returns both date and time components
//! - Based on computer's system clock
//! - No parameters required (parameterless function)
//! - Commonly used for timestamps, logging, and timing operations
//! - Can be separated into date-only or time-only using `Date()` or `Time()`
//! - Precision to the second (does not include milliseconds)
//! - Returns Variant (Date) type
//! - Subject to system time zone settings
//!
//! ### Comparison with Related Functions:
//! - **Date** - Returns only the date portion (time is midnight)
//! - **Time** - Returns only the time portion (date is December 30, 1899)
//! - **Now** - Returns both date and time components
//! - **Timer** - Returns seconds since midnight as Single (for precision timing)
//!
//! ### Common Use Cases:
//! - Create timestamps for logging
//! - Record when events occur
//! - Calculate elapsed time
//! - Display current date and time to users
//! - Set default values for date fields
//! - Generate time-based filenames
//! - Track operation start/end times
//! - Audit trail creation
//!
//! ## Typical Uses
//!
//! 1. **Timestamps** - Record when operations occur
//! 2. **Logging** - Add timestamps to log entries
//! 3. **Audit Trails** - Track when records are created/modified
//! 4. **Performance Timing** - Measure operation duration
//! 5. **Display Current Time** - Show users the current date/time
//! 6. **Default Values** - Initialize date fields with current date/time
//! 7. **File Naming** - Create time-stamped file names
//! 8. **Session Tracking** - Record login/logout times
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get current date and time
//! Dim currentDateTime As Date
//! currentDateTime = Now
//! ```
//!
//! ```vb
//! ' Example 2: Display to user
//! MsgBox "Current time is: " & Now
//! ```
//!
//! ```vb
//! ' Example 3: Create timestamp
//! Dim timestamp As String
//! timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
//! ```
//!
//! ```vb
//! ' Example 4: Calculate elapsed time
//! Dim startTime As Date
//! startTime = Now
//! ' ... do some work ...
//! MsgBox "Elapsed: " & DateDiff("s", startTime, Now) & " seconds"
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Simple timestamp logging
//! Sub LogMessage(message As String)
//!     Debug.Print Now & " - " & message
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 2: Formatted timestamp
//! Function GetTimestamp() As String
//!     GetTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 3: Calculate operation duration
//! Function MeasureOperation() As Double
//!     Dim startTime As Date
//!     Dim endTime As Date
//!     
//!     startTime = Now
//!     
//!     ' Perform operation
//!     DoSomething
//!     
//!     endTime = Now
//!     
//!     ' Return elapsed seconds
//!     MeasureOperation = DateDiff("s", startTime, endTime)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 4: Audit trail update
//! Sub UpdateRecord(recordID As Long)
//!     Dim sql As String
//!     sql = "UPDATE Records SET ModifiedDate = " & _
//!           Format(Now, "\#mm\/dd\/yyyy hh:nn:ss\#") & _
//!           " WHERE ID = " & recordID
//!     ExecuteSQL sql
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 5: Time-stamped filename
//! Function GetLogFileName() As String
//!     GetLogFileName = "log_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Session tracking
//! Sub RecordLogin(userID As Long)
//!     Dim loginTime As Date
//!     loginTime = Now
//!     
//!     ' Store in database or session object
//!     Session("LoginTime") = loginTime
//!     Session("UserID") = userID
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 7: Timeout checking
//! Function IsTimedOut(startTime As Date, timeoutMinutes As Long) As Boolean
//!     Dim elapsed As Long
//!     elapsed = DateDiff("n", startTime, Now)
//!     IsTimedOut = (elapsed >= timeoutMinutes)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: Scheduled task checking
//! Function ShouldRunTask(lastRun As Date, intervalHours As Long) As Boolean
//!     Dim hoursSinceRun As Long
//!     
//!     If IsNull(lastRun) Then
//!         ShouldRunTask = True
//!     Else
//!         hoursSinceRun = DateDiff("h", lastRun, Now)
//!         ShouldRunTask = (hoursSinceRun >= intervalHours)
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Display relative time
//! Function GetTimeAgo(pastTime As Date) As String
//!     Dim seconds As Long
//!     Dim minutes As Long
//!     Dim hours As Long
//!     
//!     seconds = DateDiff("s", pastTime, Now)
//!     
//!     If seconds < 60 Then
//!         GetTimeAgo = seconds & " seconds ago"
//!     ElseIf seconds < 3600 Then
//!         minutes = seconds \ 60
//!         GetTimeAgo = minutes & " minutes ago"
//!     Else
//!         hours = seconds \ 3600
//!         GetTimeAgo = hours & " hours ago"
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Business hours check
//! Function IsDuringBusinessHours() As Boolean
//!     Dim currentHour As Integer
//!     Dim currentDay As Integer
//!     
//!     currentHour = Hour(Now)
//!     currentDay = Weekday(Now)
//!     
//!     ' Monday-Friday, 9 AM to 5 PM
//!     IsDuringBusinessHours = (currentDay >= vbMonday And currentDay <= vbFriday) And _
//!                             (currentHour >= 9 And currentHour < 17)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Performance Monitor Class
//!
//! ```vb
//! ' Class: PerformanceMonitor
//! ' Tracks operation performance with detailed timing
//!
//! Option Explicit
//!
//! Private m_operations As Collection
//!
//! Private Type OperationTiming
//!     operationName As String
//!     startTime As Date
//!     endTime As Date
//!     duration As Double
//! End Type
//!
//! Private Sub Class_Initialize()
//!     Set m_operations = New Collection
//! End Sub
//!
//! Public Sub StartOperation(operationName As String)
//!     Dim timing As OperationTiming
//!     
//!     timing.operationName = operationName
//!     timing.startTime = Now
//!     timing.endTime = 0
//!     timing.duration = 0
//!     
//!     m_operations.Add timing, operationName
//! End Sub
//!
//! Public Sub EndOperation(operationName As String)
//!     Dim timing As OperationTiming
//!     Dim i As Long
//!     
//!     ' Find the operation
//!     For i = 1 To m_operations.Count
//!         timing = m_operations(i)
//!         If timing.operationName = operationName Then
//!             timing.endTime = Now
//!             timing.duration = DateDiff("s", timing.startTime, timing.endTime)
//!             
//!             ' Update the collection
//!             m_operations.Remove i
//!             m_operations.Add timing, operationName
//!             Exit Sub
//!         End If
//!     Next i
//! End Sub
//!
//! Public Function GetDuration(operationName As String) As Double
//!     Dim timing As OperationTiming
//!     Dim i As Long
//!     
//!     For i = 1 To m_operations.Count
//!         timing = m_operations(i)
//!         If timing.operationName = operationName Then
//!             If timing.endTime = 0 Then
//!                 ' Still running - calculate current duration
//!                 GetDuration = DateDiff("s", timing.startTime, Now)
//!             Else
//!                 GetDuration = timing.duration
//!             End If
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetDuration = -1 ' Not found
//! End Function
//!
//! Public Function GenerateReport() As String
//!     Dim report As String
//!     Dim timing As OperationTiming
//!     Dim i As Long
//!     
//!     report = "Performance Report - " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
//!     report = report & String(60, "-") & vbCrLf
//!     
//!     For i = 1 To m_operations.Count
//!         timing = m_operations(i)
//!         report = report & timing.operationName & ": "
//!         
//!         If timing.endTime = 0 Then
//!             report = report & "Running (" & DateDiff("s", timing.startTime, Now) & "s)"
//!         Else
//!             report = report & timing.duration & " seconds"
//!         End If
//!         
//!         report = report & vbCrLf
//!     Next i
//!     
//!     GenerateReport = report
//! End Function
//! ```
//!
//! ### Example 2: Audit Logger Class
//!
//! ```vb
//! ' Class: AuditLogger
//! ' Logs all database operations with timestamps
//!
//! Option Explicit
//!
//! Private m_logFile As String
//! Private m_enabled As Boolean
//!
//! Public Sub Initialize(logFilePath As String)
//!     m_logFile = logFilePath
//!     m_enabled = True
//! End Sub
//!
//! Public Sub LogInsert(tableName As String, recordID As Variant, userName As String)
//!     Dim entry As String
//!     entry = FormatLogEntry("INSERT", tableName, recordID, userName, "")
//!     WriteToLog entry
//! End Sub
//!
//! Public Sub LogUpdate(tableName As String, recordID As Variant, userName As String, changes As String)
//!     Dim entry As String
//!     entry = FormatLogEntry("UPDATE", tableName, recordID, userName, changes)
//!     WriteToLog entry
//! End Sub
//!
//! Public Sub LogDelete(tableName As String, recordID As Variant, userName As String)
//!     Dim entry As String
//!     entry = FormatLogEntry("DELETE", tableName, recordID, userName, "")
//!     WriteToLog entry
//! End Sub
//!
//! Public Sub LogSelect(tableName As String, userName As String, criteria As String)
//!     Dim entry As String
//!     entry = FormatLogEntry("SELECT", tableName, "", userName, criteria)
//!     WriteToLog entry
//! End Sub
//!
//! Private Function FormatLogEntry(operation As String, _
//!                                tableName As String, _
//!                                recordID As Variant, _
//!                                userName As String, _
//!                                details As String) As String
//!     Dim entry As String
//!     
//!     entry = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab
//!     entry = entry & operation & vbTab
//!     entry = entry & tableName & vbTab
//!     entry = entry & CStr(recordID) & vbTab
//!     entry = entry & userName & vbTab
//!     entry = entry & details
//!     
//!     FormatLogEntry = entry
//! End Function
//!
//! Private Sub WriteToLog(entry As String)
//!     Dim fileNum As Integer
//!     
//!     If Not m_enabled Then Exit Sub
//!     
//!     On Error Resume Next
//!     fileNum = FreeFile
//!     Open m_logFile For Append As #fileNum
//!     Print #fileNum, entry
//!     Close #fileNum
//!     On Error GoTo 0
//! End Sub
//!
//! Public Function GetRecentEntries(minutes As Long) As Collection
//!     Dim entries As New Collection
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim timestamp As Date
//!     Dim cutoffTime As Date
//!     
//!     cutoffTime = DateAdd("n", -minutes, Now)
//!     
//!     On Error Resume Next
//!     fileNum = FreeFile
//!     Open m_logFile For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         ' Parse timestamp from first field
//!         timestamp = CDate(Left(line, 19))
//!         
//!         If timestamp >= cutoffTime Then
//!             entries.Add line
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//!     On Error GoTo 0
//!     
//!     Set GetRecentEntries = entries
//! End Function
//! ```
//!
//! ### Example 3: Session Manager Module
//!
//! ```vb
//! ' Module: SessionManager
//! ' Manages user sessions with timeout tracking
//!
//! Option Explicit
//!
//! Private Type UserSession
//!     userID As Long
//!     userName As String
//!     loginTime As Date
//!     lastActivity As Date
//!     ipAddress As String
//!     isActive As Boolean
//! End Type
//!
//! Private m_sessions As Collection
//! Private m_timeoutMinutes As Long
//!
//! Public Sub Initialize(timeoutMinutes As Long)
//!     Set m_sessions = New Collection
//!     m_timeoutMinutes = timeoutMinutes
//! End Sub
//!
//! Public Function CreateSession(userID As Long, userName As String, ipAddress As String) As String
//!     Dim session As UserSession
//!     Dim sessionID As String
//!     
//!     ' Generate unique session ID
//!     sessionID = GenerateSessionID()
//!     
//!     session.userID = userID
//!     session.userName = userName
//!     session.loginTime = Now
//!     session.lastActivity = Now
//!     session.ipAddress = ipAddress
//!     session.isActive = True
//!     
//!     m_sessions.Add session, sessionID
//!     
//!     CreateSession = sessionID
//! End Function
//!
//! Public Sub UpdateActivity(sessionID As String)
//!     Dim session As UserSession
//!     
//!     On Error Resume Next
//!     session = m_sessions(sessionID)
//!     
//!     If Err.Number = 0 Then
//!         session.lastActivity = Now
//!         m_sessions.Remove sessionID
//!         m_sessions.Add session, sessionID
//!     End If
//!     On Error GoTo 0
//! End Sub
//!
//! Public Function IsSessionValid(sessionID As String) As Boolean
//!     Dim session As UserSession
//!     Dim minutesIdle As Long
//!     
//!     On Error Resume Next
//!     session = m_sessions(sessionID)
//!     
//!     If Err.Number <> 0 Then
//!         IsSessionValid = False
//!         Exit Function
//!     End If
//!     On Error GoTo 0
//!     
//!     If Not session.isActive Then
//!         IsSessionValid = False
//!         Exit Function
//!     End If
//!     
//!     minutesIdle = DateDiff("n", session.lastActivity, Now)
//!     IsSessionValid = (minutesIdle < m_timeoutMinutes)
//! End Function
//!
//! Public Sub CleanupExpiredSessions()
//!     Dim session As UserSession
//!     Dim sessionID As Variant
//!     Dim minutesIdle As Long
//!     Dim expiredIDs As Collection
//!     
//!     Set expiredIDs = New Collection
//!     
//!     ' Find expired sessions
//!     For Each sessionID In m_sessions
//!         session = m_sessions(sessionID)
//!         minutesIdle = DateDiff("n", session.lastActivity, Now)
//!         
//!         If minutesIdle >= m_timeoutMinutes Then
//!             expiredIDs.Add sessionID
//!         End If
//!     Next sessionID
//!     
//!     ' Remove expired sessions
//!     For Each sessionID In expiredIDs
//!         m_sessions.Remove sessionID
//!     Next sessionID
//! End Sub
//!
//! Public Function GetSessionDuration(sessionID As String) As Long
//!     Dim session As UserSession
//!     
//!     On Error Resume Next
//!     session = m_sessions(sessionID)
//!     
//!     If Err.Number = 0 Then
//!         GetSessionDuration = DateDiff("n", session.loginTime, Now)
//!     Else
//!         GetSessionDuration = 0
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! Private Function GenerateSessionID() As String
//!     GenerateSessionID = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000)
//! End Function
//! ```
//!
//! ### Example 4: Scheduled Task Runner
//!
//! ```vb
//! ' Class: ScheduledTaskRunner
//! ' Executes tasks on a schedule based on current time
//!
//! Option Explicit
//!
//! Private Type ScheduledTask
//!     taskName As String
//!     lastRun As Date
//!     intervalMinutes As Long
//!     enabled As Boolean
//!     callbackObject As Object
//!     callbackMethod As String
//! End Type
//!
//! Private m_tasks As Collection
//!
//! Private Sub Class_Initialize()
//!     Set m_tasks = New Collection
//! End Sub
//!
//! Public Sub AddTask(taskName As String, _
//!                   intervalMinutes As Long, _
//!                   callbackObj As Object, _
//!                   callbackMethod As String)
//!     Dim task As ScheduledTask
//!     
//!     task.taskName = taskName
//!     task.lastRun = 0
//!     task.intervalMinutes = intervalMinutes
//!     task.enabled = True
//!     Set task.callbackObject = callbackObj
//!     task.callbackMethod = callbackMethod
//!     
//!     m_tasks.Add task, taskName
//! End Sub
//!
//! Public Sub CheckAndRunTasks()
//!     Dim task As ScheduledTask
//!     Dim i As Long
//!     Dim minutesSinceRun As Long
//!     
//!     For i = 1 To m_tasks.Count
//!         task = m_tasks(i)
//!         
//!         If task.enabled Then
//!             If task.lastRun = 0 Then
//!                 ' Never run - execute now
//!                 ExecuteTask task
//!                 task.lastRun = Now
//!                 UpdateTask i, task
//!             Else
//!                 minutesSinceRun = DateDiff("n", task.lastRun, Now)
//!                 
//!                 If minutesSinceRun >= task.intervalMinutes Then
//!                     ExecuteTask task
//!                     task.lastRun = Now
//!                     UpdateTask i, task
//!                 End If
//!             End If
//!         End If
//!     Next i
//! End Sub
//!
//! Private Sub ExecuteTask(task As ScheduledTask)
//!     On Error Resume Next
//!     CallByName task.callbackObject, task.callbackMethod, VbMethod
//!     
//!     If Err.Number <> 0 Then
//!         Debug.Print "Task execution error: " & task.taskName & " - " & Err.Description
//!     End If
//!     On Error GoTo 0
//! End Sub
//!
//! Private Sub UpdateTask(index As Long, task As ScheduledTask)
//!     Dim taskName As String
//!     taskName = task.taskName
//!     
//!     m_tasks.Remove index
//!     m_tasks.Add task, taskName
//! End Sub
//!
//! Public Function GetNextRunTime(taskName As String) As Date
//!     Dim task As ScheduledTask
//!     
//!     On Error Resume Next
//!     task = m_tasks(taskName)
//!     
//!     If Err.Number = 0 Then
//!         If task.lastRun = 0 Then
//!             GetNextRunTime = Now ' Will run immediately
//!         Else
//!             GetNextRunTime = DateAdd("n", task.intervalMinutes, task.lastRun)
//!         End If
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! Public Sub EnableTask(taskName As String)
//!     Dim task As ScheduledTask
//!     Dim i As Long
//!     
//!     For i = 1 To m_tasks.Count
//!         task = m_tasks(i)
//!         If task.taskName = taskName Then
//!             task.enabled = True
//!             UpdateTask i, task
//!             Exit Sub
//!         End If
//!     Next i
//! End Sub
//!
//! Public Sub DisableTask(taskName As String)
//!     Dim task As ScheduledTask
//!     Dim i As Long
//!     
//!     For i = 1 To m_tasks.Count
//!         task = m_tasks(i)
//!         If task.taskName = taskName Then
//!             task.enabled = False
//!             UpdateTask i, task
//!             Exit Sub
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Now rarely fails, but system clock issues can occur:
//! On Error Resume Next
//! Dim currentTime As Date
//! currentTime = Now
//! If Err.Number <> 0 Then
//!     MsgBox "Unable to get system time: " & Err.Description
//!     ' Use a default or cached value
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - Now is a fast function - safe to call frequently
//! - For high-precision timing, use Timer function instead
//! - Does not include milliseconds - precision is to the second
//! - Caching Now value in tight loops can improve performance slightly
//! - System clock access is generally very fast
//! - No performance difference between Now, Date, and Time functions
//!
//! ## Best Practices
//!
//! 1. **Use for timestamps** - Ideal for logging and audit trails
//! 2. **Store in Date variables** - Declare as Date type, not Variant when possible
//! 3. **Format for display** - Use `Format()` function for user-friendly output
//! 4. **Consider time zones** - Now uses local system time, not UTC
//! 5. **Use Timer for precision** - For sub-second timing, use Timer function
//! 6. **Cache in loops** - Store Now once at loop start instead of calling repeatedly
//! 7. **Document timezone** - Make it clear whether times are local or UTC
//! 8. **Use `DateDiff` carefully** - Be aware of daylight saving time changes
//! 9. **Validate before math** - Check for valid dates before date arithmetic
//! 10. **Consider Date vs Now** - Use Date if you only need the date portion
//!
//! ## Comparison with Alternatives
//!
//! | Function | Date Component | Time Component | Use Case |
//! |----------|---------------|----------------|----------|
//! | **Now** | Yes (current) | Yes (current) | Full timestamp |
//! | **Date** | Yes (current) | No (midnight) | Date-only operations |
//! | **Time** | No (12/30/1899) | Yes (current) | Time-only operations |
//! | **Timer** | No | Yes (as Single) | High-precision timing |
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, Word, etc.)
//! - Available in VB6
//! - Available in `VBScript`
//! - Uses Windows system clock
//! - Subject to system time zone settings
//! - Affected by daylight saving time changes
//! - No parameters required (parameterless function)
//! - Returns local time, not UTC
//!
//! ## Limitations
//!
//! - Returns local time only (no UTC option)
//! - Precision limited to seconds (no milliseconds)
//! - Subject to system clock changes
//! - Date range: January 1, 100 to December 31, 9999
//! - Can be affected by daylight saving time transitions
//! - No built-in time zone conversion
//! - Depends on accurate system clock
//!
//! ## Related Functions
//!
//! - **Date** - Returns current date (time = midnight)
//! - **Time** - Returns current time (date = 12/30/1899)
//! - **Timer** - Returns seconds since midnight (for precision timing)
//! - **`DateAdd`** - Adds time intervals to dates
//! - **`DateDiff`** - Calculates difference between dates
//! - **Format** - Formats date/time for display
//!
//! ## VB6 Parser Notes
//!
//! Now is a parameterless function that is parsed as a `CallExpression`. This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers working
//! with date and time operations, timestamps, logging, and time-based calculations.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn now_basic() {
        let source = r"
Dim currentTime As Date
currentTime = Now
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_parentheses() {
        let source = r"
Dim dt As Date
dt = Now()
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_if_statement() {
        let source = r#"
If Now > deadline Then
    MsgBox "Overdue"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_function_return() {
        let source = r"
Function GetCurrentTime() As Date
    GetCurrentTime = Now
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_concatenation() {
        let source = r#"
Dim msg As String
msg = "Current time: " & Now
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_debug_print() {
        let source = r#"
Debug.Print "Timestamp: " & Now
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_msgbox() {
        let source = r#"
MsgBox "Current time is: " & Now
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_format() {
        let source = r#"
Dim formatted As String
formatted = Format(Now, "yyyy-mm-dd hh:nn:ss")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_datediff() {
        let source = r#"
Dim elapsed As Long
elapsed = DateDiff("s", startTime, Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_class_usage() {
        let source = r"
Private m_timestamp As Date

Public Sub UpdateTimestamp()
    m_timestamp = Now
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_with_statement() {
        let source = r"
With currentRecord
    .CreatedDate = Now
    .ModifiedDate = Now
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_array_assignment() {
        let source = r"
Dim timestamps(10) As Date
timestamps(i) = Now
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_property_assignment() {
        let source = r"
Set obj = New Logger
obj.Timestamp = Now
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_select_case() {
        let source = r#"
Select Case Hour(Now)
    Case 0 To 11
        greeting = "Good morning"
    Case 12 To 17
        greeting = "Good afternoon"
    Case Else
        greeting = "Good evening"
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_elseif() {
        let source = r"
If x > 0 Then
    y = 1
ElseIf Now > deadline Then
    y = 2
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_for_loop() {
        let source = r#"
Dim startTime As Date
startTime = Now
For i = 1 To 1000
    DoWork
Next i
MsgBox "Elapsed: " & DateDiff("s", startTime, Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_do_while() {
        let source = r"
Do While Now < endTime
    ProcessData
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_do_until() {
        let source = r"
Do Until Now >= targetTime
    WaitForEvent
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_while_wend() {
        let source = r"
While Now < cutoffTime
    count = count + 1
Wend
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_iif() {
        let source = r#"
Dim status As String
status = IIf(Now > deadline, "Late", "On time")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_comparison() {
        let source = r#"
If DateDiff("h", lastUpdate, Now) > 24 Then
    UpdateData
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_function_argument() {
        let source = r#"
Call LogEvent("User login", Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_sql_insert() {
        let source = r#"
sql = "INSERT INTO Events (Timestamp) VALUES (" & Format(Now, "\#mm\/dd\/yyyy hh:nn:ss\#") & ")"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_year_function() {
        let source = r"
Dim currentYear As Integer
currentYear = Year(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_month_function() {
        let source = r"
Dim currentMonth As Integer
currentMonth = Month(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_dateadd() {
        let source = r#"
Dim futureDate As Date
futureDate = DateAdd("d", 7, Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn now_multiple_calls() {
        let source = r"
Dim start As Date
Dim finish As Date
start = Now
DoWork
finish = Now
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/datetime/now");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
