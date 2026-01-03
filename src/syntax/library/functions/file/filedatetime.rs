//! # `FileDateTime` Function
//!
//! Returns a `Variant` (`Date`) representing the date and time when a file was created or last modified.
//!
//! ## Syntax
//!
//! ```vb
//! FileDateTime(pathname)
//! ```
//!
//! ## Parameters
//!
//! - **pathname**: Required. A `String` expression that specifies a file name. May include
//!   directory or folder, and drive. If the file is not found, an error occurs.
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `Date` representing the date and time when the file was
//! last modified. The returned value includes both date and time components.
//!
//! ## Remarks
//!
//! The `FileDateTime` function returns the last modification date and time of a file.
//! This is useful for file management, backup utilities, synchronization, and determining
//! if files need to be updated.
//!
//! **Important Characteristics:**
//!
//! - Returns date/time of last modification
//! - File does not need to be open
//! - Error if file does not exist (Error 53)
//! - Error if path is invalid (Error 76)
//! - Works with full paths and relative paths
//! - Returns same value as file system shows
//! - Precision depends on file system (typically 2-second resolution on FAT, 100ns on NTFS)
//! - Affected by daylight saving time changes
//! - Returns local time (not UTC)
//! - Can be used with wildcards via `Dir` function
//!
//! ## Typical Uses
//!
//! - Determine if file is newer than another
//! - Check if file has been modified since last access
//! - File synchronization between locations
//! - Backup and archive utilities
//! - Change detection systems
//! - File age calculations
//! - Automated cleanup of old files
//! - Build systems (check if recompilation needed)
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim fileDate As Date
//!
//! ' Get file modification date/time
//! fileDate = FileDateTime("C:\data.txt")
//! Debug.Print "File last modified: " & fileDate
//!
//! ' Format the date/time
//! Debug.Print Format(fileDate, "yyyy-mm-dd hh:nn:ss")
//! ```
//!
//! ### Compare File Dates
//!
//! ```vb
//! Function IsFileNewer(file1 As String, file2 As String) As Boolean
//!     ' Returns True if file1 is newer than file2
//!     On Error GoTo ErrorHandler
//!     
//!     Dim date1 As Date
//!     Dim date2 As Date
//!     
//!     date1 = FileDateTime(file1)
//!     date2 = FileDateTime(file2)
//!     
//!     IsFileNewer = (date1 > date2)
//!     Exit Function
//!     
//! ErrorHandler:
//!     IsFileNewer = False
//! End Function
//! ```
//!
//! ### Check If File Was Modified Today
//!
//! ```vb
//! Function IsModifiedToday(filePath As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileDate As Date
//!     fileDate = FileDateTime(filePath)
//!     
//!     ' Compare just the date part
//!     IsModifiedToday = (Int(fileDate) = Int(Date))
//!     Exit Function
//!     
//! ErrorHandler:
//!     IsModifiedToday = False
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Find Newest File in Directory
//!
//! ```vb
//! Function FindNewestFile(folderPath As String, pattern As String) As String
//!     Dim fileName As String
//!     Dim newestFile As String
//!     Dim newestDate As Date
//!     Dim currentDate As Date
//!     
//!     ' Ensure path ends with backslash
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     fileName = Dir(folderPath & pattern)
//!     newestDate = #1/1/1900#
//!     
//!     Do While fileName <> ""
//!         On Error Resume Next
//!         currentDate = FileDateTime(folderPath & fileName)
//!         
//!         If Err.Number = 0 Then
//!             If currentDate > newestDate Then
//!                 newestDate = currentDate
//!                 newestFile = fileName
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     FindNewestFile = newestFile
//! End Function
//! ```
//!
//! ### Get File Age
//!
//! ```vb
//! Function GetFileAgeInDays(filePath As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileDate As Date
//!     fileDate = FileDateTime(filePath)
//!     
//!     ' Calculate difference in days
//!     GetFileAgeInDays = DateDiff("d", fileDate, Now)
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileAgeInDays = -1
//! End Function
//!
//! Function GetFileAgeInHours(filePath As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileDate As Date
//!     fileDate = FileDateTime(filePath)
//!     
//!     GetFileAgeInHours = DateDiff("h", fileDate, Now)
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetFileAgeInHours = -1
//! End Function
//! ```
//!
//! ### Delete Old Files
//!
//! ```vb
//! Sub DeleteOldFiles(folderPath As String, daysOld As Long)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileDate As Date
//!     Dim ageInDays As Long
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
//!         fileDate = FileDateTime(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             ageInDays = DateDiff("d", fileDate, Now)
//!             
//!             If ageInDays > daysOld Then
//!                 Kill fullPath
//!                 Debug.Print "Deleted: " & fileName & " (Age: " & ageInDays & " days)"
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//! End Sub
//! ```
//!
//! ### Check If File Needs Backup
//!
//! ```vb
//! Function NeedsBackup(sourceFile As String, backupFile As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim sourceDate As Date
//!     Dim backupDate As Date
//!     
//!     ' Get source file date
//!     sourceDate = FileDateTime(sourceFile)
//!     
//!     ' Check if backup exists
//!     If Dir(backupFile) = "" Then
//!         ' No backup exists
//!         NeedsBackup = True
//!         Exit Function
//!     End If
//!     
//!     ' Get backup file date
//!     backupDate = FileDateTime(backupFile)
//!     
//!     ' Needs backup if source is newer than backup
//!     NeedsBackup = (sourceDate > backupDate)
//!     Exit Function
//!     
//! ErrorHandler:
//!     NeedsBackup = True  ' Assume needs backup on error
//! End Function
//! ```
//!
//! ### List Files Modified Within Date Range
//!
//! ```vb
//! Function GetFilesModifiedBetween(folderPath As String, startDate As Date, _
//!                                   endDate As Date) As Collection
//!     Dim files As New Collection
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileDate As Date
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
//!         fileDate = FileDateTime(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             If fileDate >= startDate And fileDate <= endDate Then
//!                 files.Add fullPath
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     Set GetFilesModifiedBetween = files
//! End Function
//! ```
//!
//! ### File Synchronization Check
//!
//! ```vb
//! Function SynchronizeFile(sourceFile As String, destFile As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim sourceDate As Date
//!     Dim destDate As Date
//!     Dim needsCopy As Boolean
//!     
//!     sourceDate = FileDateTime(sourceFile)
//!     
//!     ' Check if destination exists
//!     If Dir(destFile) = "" Then
//!         needsCopy = True
//!     Else
//!         destDate = FileDateTime(destFile)
//!         needsCopy = (sourceDate > destDate)
//!     End If
//!     
//!     If needsCopy Then
//!         FileCopy sourceFile, destFile
//!         Debug.Print "Synchronized: " & sourceFile & " -> " & destFile
//!         SynchronizeFile = True
//!     Else
//!         Debug.Print "Already synchronized: " & destFile
//!         SynchronizeFile = False
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Debug.Print "Error synchronizing: " & Err.Description
//!     SynchronizeFile = False
//! End Function
//! ```
//!
//! ### Build System - Check Dependencies
//!
//! ```vb
//! Function SourceNewerThanExecutable(sourceFile As String, exeFile As String) As Boolean
//!     ' Used in build systems to determine if recompilation is needed
//!     On Error GoTo ErrorHandler
//!     
//!     Dim sourceDate As Date
//!     Dim exeDate As Date
//!     
//!     ' Check if executable exists
//!     If Dir(exeFile) = "" Then
//!         SourceNewerThanExecutable = True  ' Need to build
//!         Exit Function
//!     End If
//!     
//!     sourceDate = FileDateTime(sourceFile)
//!     exeDate = FileDateTime(exeFile)
//!     
//!     SourceNewerThanExecutable = (sourceDate > exeDate)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SourceNewerThanExecutable = True  ' Assume needs build on error
//! End Function
//! ```
//!
//! ### Generate File Report
//!
//! ```vb
//! Sub GenerateFileReport(folderPath As String, reportFile As String)
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim fileDate As Date
//!     Dim reportNum As Integer
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     reportNum = FreeFile
//!     Open reportFile For Output As #reportNum
//!     
//!     Print #reportNum, "File Report for: " & folderPath
//!     Print #reportNum, "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
//!     Print #reportNum, String(80, "-")
//!     Print #reportNum, "Filename", "Modified Date", "Age (Days)"
//!     Print #reportNum, String(80, "-")
//!     
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         fileDate = FileDateTime(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             Print #reportNum, fileName, _
//!                   Format(fileDate, "yyyy-mm-dd hh:nn:ss"), _
//!                   DateDiff("d", fileDate, Now)
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     Close #reportNum
//! End Sub
//! ```
//!
//! ### Monitor File Changes
//!
//! ```vb
//! Private lastCheckedDate As Date
//!
//! Function HasFileChanged(filePath As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim currentDate As Date
//!     currentDate = FileDateTime(filePath)
//!     
//!     If lastCheckedDate = 0 Then
//!         ' First check
//!         lastCheckedDate = currentDate
//!         HasFileChanged = False
//!     Else
//!         HasFileChanged = (currentDate > lastCheckedDate)
//!         If HasFileChanged Then
//!             lastCheckedDate = currentDate
//!         End If
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     HasFileChanged = False
//! End Function
//! ```
//!
//! ### Sort Files by Date
//!
//! ```vb
//! Type FileInfo
//!     Name As String
//!     ModifiedDate As Date
//! End Type
//!
//! Function GetFilesSortedByDate(folderPath As String) As Variant
//!     Dim files() As FileInfo
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim count As Long
//!     Dim i As Long, j As Long
//!     Dim temp As FileInfo
//!     
//!     If Right(folderPath, 1) <> "\" Then
//!         folderPath = folderPath & "\"
//!     End If
//!     
//!     ' Collect files
//!     ReDim files(0 To 100)
//!     count = 0
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         files(count).Name = fileName
//!         files(count).ModifiedDate = FileDateTime(fullPath)
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
//!         ' Bubble sort by date (newest first)
//!         For i = 0 To count - 2
//!             For j = i + 1 To count - 1
//!                 If files(j).ModifiedDate > files(i).ModifiedDate Then
//!                     temp = files(i)
//!                     files(i) = files(j)
//!                     files(j) = temp
//!                 End If
//!             Next j
//!         Next i
//!     End If
//!     
//!     GetFilesSortedByDate = files
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Incremental Backup System
//!
//! ```vb
//! Function PerformIncrementalBackup(sourceFolder As String, backupFolder As String, _
//!                                    lastBackupDate As Date) As Long
//!     Dim fileName As String
//!     Dim sourcePath As String
//!     Dim backupPath As String
//!     Dim fileDate As Date
//!     Dim copiedCount As Long
//!     
//!     If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
//!     If Right(backupFolder, 1) <> "\" Then backupFolder = backupFolder & "\"
//!     
//!     ' Ensure backup folder exists
//!     If Dir(backupFolder, vbDirectory) = "" Then
//!         MkDir backupFolder
//!     End If
//!     
//!     fileName = Dir(sourceFolder & "*.*")
//!     copiedCount = 0
//!     
//!     Do While fileName <> ""
//!         sourcePath = sourceFolder & fileName
//!         backupPath = backupFolder & fileName
//!         
//!         On Error Resume Next
//!         fileDate = FileDateTime(sourcePath)
//!         
//!         If Err.Number = 0 And fileDate > lastBackupDate Then
//!             FileCopy sourcePath, backupPath
//!             If Err.Number = 0 Then
//!                 copiedCount = copiedCount + 1
//!                 Debug.Print "Backed up: " & fileName
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     PerformIncrementalBackup = copiedCount
//! End Function
//! ```
//!
//! ### File Cache Invalidation
//!
//! ```vb
//! Private Type CacheEntry
//!     FilePath As String
//!     CachedDate As Date
//!     CachedData As Variant
//! End Type
//!
//! Private cache() As CacheEntry
//! Private cacheCount As Long
//!
//! Function GetCachedFileData(filePath As String) As Variant
//!     Dim i As Long
//!     Dim currentDate As Date
//!     
//!     On Error Resume Next
//!     currentDate = FileDateTime(filePath)
//!     
//!     If Err.Number <> 0 Then
//!         GetCachedFileData = Null
//!         Exit Function
//!     End If
//!     
//!     ' Check if in cache and still valid
//!     For i = 0 To cacheCount - 1
//!         If cache(i).FilePath = filePath Then
//!             If cache(i).CachedDate = currentDate Then
//!                 ' Cache is valid
//!                 GetCachedFileData = cache(i).CachedData
//!                 Exit Function
//!             Else
//!                 ' Cache is stale, reload
//!                 cache(i).CachedData = LoadFileData(filePath)
//!                 cache(i).CachedDate = currentDate
//!                 GetCachedFileData = cache(i).CachedData
//!                 Exit Function
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Not in cache, add it
//!     ReDim Preserve cache(0 To cacheCount)
//!     cache(cacheCount).FilePath = filePath
//!     cache(cacheCount).CachedDate = currentDate
//!     cache(cacheCount).CachedData = LoadFileData(filePath)
//!     GetCachedFileData = cache(cacheCount).CachedData
//!     cacheCount = cacheCount + 1
//! End Function
//! ```
//!
//! ### Multi-Directory Synchronization
//!
//! ```vb
//! Function SynchronizeFolders(sourceFolder As String, destFolder As String) As Long
//!     Dim fileName As String
//!     Dim sourcePath As String
//!     Dim destPath As String
//!     Dim sourceDate As Date
//!     Dim destDate As Date
//!     Dim syncCount As Long
//!     
//!     If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
//!     If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"
//!     
//!     ' Ensure destination exists
//!     If Dir(destFolder, vbDirectory) = "" Then
//!         MkDir destFolder
//!     End If
//!     
//!     fileName = Dir(sourceFolder & "*.*")
//!     syncCount = 0
//!     
//!     Do While fileName <> ""
//!         sourcePath = sourceFolder & fileName
//!         destPath = destFolder & fileName
//!         
//!         On Error Resume Next
//!         sourceDate = FileDateTime(sourcePath)
//!         
//!         If Err.Number = 0 Then
//!             ' Check if destination exists
//!             If Dir(destPath) = "" Then
//!                 ' Destination doesn't exist, copy
//!                 FileCopy sourcePath, destPath
//!                 syncCount = syncCount + 1
//!             Else
//!                 ' Check dates
//!                 destDate = FileDateTime(destPath)
//!                 If sourceDate > destDate Then
//!                     ' Source is newer, update
//!                     FileCopy sourcePath, destPath
//!                     syncCount = syncCount + 1
//!                 End If
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     SynchronizeFolders = syncCount
//! End Function
//! ```
//!
//! ### Change Detection System
//!
//! ```vb
//! Type ChangeRecord
//!     FilePath As String
//!     PreviousDate As Date
//!     CurrentDate As Date
//!     ChangeType As String  ' "Modified", "Created", "Deleted"
//! End Type
//!
//! Function DetectChanges(folderPath As String, baseline() As FileInfo) As Variant
//!     Dim changes() As ChangeRecord
//!     Dim changeCount As Long
//!     Dim fileName As String
//!     Dim fullPath As String
//!     Dim currentDate As Date
//!     Dim i As Long
//!     Dim found As Boolean
//!     
//!     If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
//!     
//!     ReDim changes(0 To 100)
//!     changeCount = 0
//!     
//!     ' Check current files
//!     fileName = Dir(folderPath & "*.*")
//!     
//!     Do While fileName <> ""
//!         fullPath = folderPath & fileName
//!         
//!         On Error Resume Next
//!         currentDate = FileDateTime(fullPath)
//!         
//!         If Err.Number = 0 Then
//!             found = False
//!             
//!             ' Check if in baseline
//!             For i = LBound(baseline) To UBound(baseline)
//!                 If baseline(i).Name = fileName Then
//!                     found = True
//!                     
//!                     ' Check if modified
//!                     If currentDate > baseline(i).ModifiedDate Then
//!                         changes(changeCount).FilePath = fullPath
//!                         changes(changeCount).PreviousDate = baseline(i).ModifiedDate
//!                         changes(changeCount).CurrentDate = currentDate
//!                         changes(changeCount).ChangeType = "Modified"
//!                         changeCount = changeCount + 1
//!                     End If
//!                     
//!                     Exit For
//!                 End If
//!             Next i
//!             
//!             ' If not found in baseline, it's new
//!             If Not found Then
//!                 changes(changeCount).FilePath = fullPath
//!                 changes(changeCount).CurrentDate = currentDate
//!                 changes(changeCount).ChangeType = "Created"
//!                 changeCount = changeCount + 1
//!             End If
//!         End If
//!         
//!         Err.Clear
//!         fileName = Dir
//!     Loop
//!     
//!     If changeCount > 0 Then
//!         ReDim Preserve changes(0 To changeCount - 1)
//!         DetectChanges = changes
//!     Else
//!         DetectChanges = Array()
//!     End If
//! End Function
//! ```
//!
//! ### Log File Rotation Based on Date
//!
//! ```vb
//! Sub RotateLogFile(logPath As String, maxAgeDays As Long)
//!     Dim logDate As Date
//!     Dim ageDays As Long
//!     Dim archivePath As String
//!     
//!     On Error Resume Next
//!     
//!     ' Check if log file exists
//!     If Dir(logPath) = "" Then Exit Sub
//!     
//!     logDate = FileDateTime(logPath)
//!     ageDays = DateDiff("d", logDate, Now)
//!     
//!     If ageDays >= maxAgeDays Then
//!         ' Create archive name with date
//!         archivePath = Left(logPath, Len(logPath) - 4) & "_" & _
//!                       Format(logDate, "yyyymmdd") & ".log"
//!         
//!         ' Rename current log to archive
//!         Name logPath As archivePath
//!         
//!         Debug.Print "Rotated log file: " & logPath & " -> " & archivePath
//!     End If
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFileDateTime(filePath As String) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     SafeFileDateTime = FileDateTime(filePath)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 53  ' File not found
//!             MsgBox "File not found: " & filePath, vbExclamation
//!             SafeFileDateTime = Null
//!         Case 76  ' Path not found
//!             MsgBox "Path not found: " & filePath, vbExclamation
//!             SafeFileDateTime = Null
//!         Case Else
//!             MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
//!             SafeFileDateTime = Null
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 53** (File not found): The specified file does not exist
//! - **Error 76** (Path not found): The specified path is invalid
//!
//! ## Performance Considerations
//!
//! - `FileDateTime` is relatively fast (reads file metadata only)
//! - Does not open the file or read contents
//! - Performance depends on file system and disk speed
//! - Network paths are slower than local paths
//! - Consider caching results if checking same file repeatedly
//! - Use `Dir` function to check existence before calling `FileDateTime`
//!
//! ## Best Practices
//!
//! ### Check File Existence First
//!
//! ```vb
//! ' Good - Check existence to avoid error
//! If Dir(filePath) <> "" Then
//!     fileDate = FileDateTime(filePath)
//! Else
//!     MsgBox "File not found"
//! End If
//!
//! ' Or use error handling
//! On Error Resume Next
//! fileDate = FileDateTime(filePath)
//! If Err.Number <> 0 Then
//!     MsgBox "Cannot get file date"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Use Full Paths
//!
//! ```vb
//! ' Good - Use full path for clarity
//! fileDate = FileDateTime("C:\Projects\data.txt")
//!
//! ' Or build from App.Path
//! fileDate = FileDateTime(App.Path & "\config.ini")
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FileDateTime` vs `Now`
//!
//! ```vb
//! ' FileDateTime - Gets file modification date
//! fileDate = FileDateTime("C:\file.txt")
//!
//! ' Now - Gets current system date/time
//! currentDate = Now
//! ```
//!
//! ### `FileDateTime` vs `GetAttr`
//!
//! ```vb
//! ' FileDateTime - Returns date/time of modification
//! fileDate = FileDateTime("C:\file.txt")
//!
//! ' GetAttr - Returns file attributes (readonly, hidden, etc.)
//! attrs = GetAttr("C:\file.txt")
//! ```
//!
//! ## Limitations
//!
//! - Returns modification date only (not creation or access date)
//! - File must exist (cannot get date of deleted files)
//! - Returns local time (not UTC)
//! - Precision limited by file system
//! - Cannot set file date/time (read-only function)
//! - Does not work with directories (use `Dir` with `vbDirectory`)
//! - No built-in wildcard support (must use with `Dir`)
//!
//! ## Related Functions
//!
//! - `Dir`: Returns file names matching a pattern
//! - `GetAttr`: Returns file attributes
//! - `FileLen`: Returns file size in bytes
//! - `Now`: Returns current system date and time
//! - `Date`: Returns current system date
//! - `DateDiff`: Calculates difference between two dates
//! - `Format`: Formats date/time for display

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn filedatetime_basic() {
        let source = r#"
fileDate = FileDateTime("C:\data.txt")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("fileDate"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"C:\\data.txt\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_variable() {
        let source = r"
fileDate = FileDateTime(filePath)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("fileDate"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("filePath"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_comparison() {
        let source = r"
isNewer = (FileDateTime(file1) > FileDateTime(file2))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("isNewer"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                ParenthesizedExpression {
                    LeftParenthesis,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("file1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        GreaterThanOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("file2"),
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
        ]);
    }

    #[test]
    fn filedatetime_debug_print() {
        let source = r#"
Debug.Print FileDateTime("C:\temp.dat")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("Debug"),
                PeriodOperator,
                PrintKeyword,
                Whitespace,
                Identifier ("FileDateTime"),
                LeftParenthesis,
                StringLiteral ("\"C:\\temp.dat\""),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_in_function() {
        let source = r#"
Function GetFileAge(path As String) As Long
    GetFileAge = DateDiff("d", FileDateTime(path), Now)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetFileAge"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("path"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                            Identifier ("GetFileAge"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("DateDiff"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"d\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("FileDateTime"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("path"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
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
    fn filedatetime_format() {
        let source = r#"
formatted = Format(FileDateTime(filePath), "yyyy-mm-dd hh:nn:ss")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("formatted"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Format"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("FileDateTime"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("filePath"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"yyyy-mm-dd hh:nn:ss\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_if_statement() {
        let source = r"
If FileDateTime(sourceFile) > FileDateTime(backupFile) Then
    needsBackup = True
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("FileDateTime"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("sourceFile"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("FileDateTime"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("backupFile"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("needsBackup"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_datediff() {
        let source = r#"
ageInDays = DateDiff("d", FileDateTime(filePath), Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("ageInDays"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DateDiff"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"d\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("FileDateTime"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("filePath"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("Now"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_concatenation() {
        let source = r#"
msg = "File modified: " & FileDateTime(filePath)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("msg"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    StringLiteralExpression {
                        StringLiteral ("\"File modified: \""),
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("FileDateTime"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("filePath"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_error_handling() {
        let source = r#"
On Error Resume Next
dt = FileDateTime(filePath)
If Err.Number <> 0 Then
    MsgBox "File not found"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                ResumeKeyword,
                Whitespace,
                NextKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("dt"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("filePath"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            IfStatement {
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
                        StringLiteral ("\"File not found\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_loop() {
        let source = r#"
Do While fileName <> ""
    currentDate = FileDateTime(folderPath & fileName)
    fileName = Dir
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("fileName"),
                    },
                    Whitespace,
                    InequalityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("currentDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("folderPath"),
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("fileName"),
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Dir"),
                        },
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_udt_field() {
        let source = r"
info.ModifiedDate = FileDateTime(filePath)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("info"),
                    PeriodOperator,
                    Identifier ("ModifiedDate"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("filePath"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_array() {
        let source = r"
dates(i) = FileDateTime(files(i))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                CallExpression {
                    Identifier ("dates"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("files"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("i"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_select_case() {
        let source = r#"
Select Case DateDiff("d", FileDateTime(filePath), Now)
    Case Is > 30
        Debug.Print "Old file"
    Case Else
        Debug.Print "Recent file"
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("DateDiff"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"d\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("FileDateTime"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("filePath"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("Now"),
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
                    IsKeyword,
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    IntegerLiteral ("30"),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("Debug"),
                            PeriodOperator,
                            PrintKeyword,
                            Whitespace,
                            StringLiteral ("\"Old file\""),
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
                            Identifier ("Debug"),
                            PeriodOperator,
                            PrintKeyword,
                            Whitespace,
                            StringLiteral ("\"Recent file\""),
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                SelectKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_msgbox() {
        let source = r#"
MsgBox "Last modified: " & FileDateTime(filePath)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Last modified: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("FileDateTime"),
                LeftParenthesis,
                Identifier ("filePath"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_max_comparison() {
        let source = r"
If FileDateTime(fullPath) > newestDate Then
    newestDate = FileDateTime(fullPath)
    newestFile = fileName
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("FileDateTime"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("fullPath"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("newestDate"),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("newestDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fullPath"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("newestFile"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("fileName"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_date_range() {
        let source = r"
If FileDateTime(fullPath) >= startDate And FileDateTime(fullPath) <= endDate Then
    files.Add fullPath
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fullPath"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        GreaterThanOrEqualOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("startDate"),
                        },
                    },
                    Whitespace,
                    AndKeyword,
                    Whitespace,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fullPath"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        LessThanOrEqualOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("endDate"),
                        },
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("files"),
                        PeriodOperator,
                        Identifier ("Add"),
                        Whitespace,
                        Identifier ("fullPath"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_print_statement() {
        let source = r"
Print #reportNum, fileName, FileDateTime(fullPath)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PrintStatement {
                PrintKeyword,
                Whitespace,
                Octothorpe,
                Identifier ("reportNum"),
                Comma,
                Whitespace,
                Identifier ("fileName"),
                Comma,
                Whitespace,
                Identifier ("FileDateTime"),
                LeftParenthesis,
                Identifier ("fullPath"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_with_app_path() {
        let source = r#"
configDate = FileDateTime(App.Path & "\config.ini")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("configDate"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("FileDateTime"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                MemberAccessExpression {
                                    Identifier ("App"),
                                    PeriodOperator,
                                    Identifier ("Path"),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\"\\config.ini\""),
                                },
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_int_function() {
        let source = r"
isToday = (Int(FileDateTime(filePath)) = Int(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("isToday"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                ParenthesizedExpression {
                    LeftParenthesis,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("Int"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("FileDateTime"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("filePath"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Int"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        DateKeyword,
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
        ]);
    }

    #[test]
    fn filedatetime_sort_comparison() {
        let source = r"
If files(j).ModifiedDate > files(i).ModifiedDate Then
    temp = files(i)
    files(i) = files(j)
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    MemberAccessExpression {
                        CallExpression {
                            Identifier ("files"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("j"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        PeriodOperator,
                        Identifier ("ModifiedDate"),
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    MemberAccessExpression {
                        CallExpression {
                            Identifier ("files"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        PeriodOperator,
                        Identifier ("ModifiedDate"),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("temp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("files"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("files"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("files"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("j"),
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
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_for_loop() {
        let source = r"
For i = 0 To fileCount - 1
    dt = FileDateTime(fileList(i))
    Debug.Print fileList(i), dt
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForStatement {
                ForKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("i"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("0"),
                },
                Whitespace,
                ToKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("fileCount"),
                    },
                    Whitespace,
                    SubtractionOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("1"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dt"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("fileList"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("i"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("fileList"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("dt"),
                        Newline,
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_multiline() {
        let source = r#"
info = "File: " & filePath & vbCrLf & _
       "Modified: " & FileDateTime(filePath) & vbCrLf & _
       "Age: " & DateDiff("d", FileDateTime(filePath), Now) & " days"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("info"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        BinaryExpression {
                            BinaryExpression {
                                BinaryExpression {
                                    BinaryExpression {
                                        BinaryExpression {
                                            BinaryExpression {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"File: \""),
                                                },
                                                Whitespace,
                                                Ampersand,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("filePath"),
                                                },
                                            },
                                            Whitespace,
                                            Ampersand,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("vbCrLf"),
                                            },
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        Underscore,
                                        Newline,
                                        Whitespace,
                                        StringLiteralExpression {
                                            StringLiteral ("\"Modified: \""),
                                        },
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("FileDateTime"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("filePath"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("vbCrLf"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            Underscore,
                            Newline,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Age: \""),
                            },
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        CallExpression {
                            Identifier ("DateDiff"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"d\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("FileDateTime"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("filePath"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\" days\""),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_cache_check() {
        let source = r"
If cache(i).CachedDate = FileDateTime(filePath) Then
    isValid = True
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    MemberAccessExpression {
                        CallExpression {
                            Identifier ("cache"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        PeriodOperator,
                        Identifier ("CachedDate"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("FileDateTime"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("filePath"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("isValid"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filedatetime_backup_check() {
        let source = r"
needsCopy = (FileDateTime(sourceFile) > FileDateTime(destFile))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("needsCopy"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                ParenthesizedExpression {
                    LeftParenthesis,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("sourceFile"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        GreaterThanOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("FileDateTime"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("destFile"),
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
        ]);
    }
}
