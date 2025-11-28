//! VB6 `Timer` Function
//!
//! The `Timer` function returns a Single representing the number of seconds that have elapsed since midnight.
//!
//! ## Syntax
//! ```vb6
//! Timer()
//! ```
//! or
//! ```vb6
//! Timer
//! ```
//!
//! ## Parameters
//! None. The `Timer` function takes no arguments.
//!
//! ## Returns
//! Returns a `Single` representing the number of seconds elapsed since midnight (00:00:00). The value ranges from 0 to 86,400 (the number of seconds in 24 hours).
//!
//! ## Remarks
//! The `Timer` function provides high-precision time measurements:
//!
//! - **No arguments**: Called without parentheses or with empty parentheses
//! - **Seconds since midnight**: Returns elapsed seconds from 00:00:00
//! - **High precision**: Resolution approximately 10-55 milliseconds depending on platform
//! - **Single type**: Returns floating-point value with fractional seconds
//! - **Midnight rollover**: Value resets to 0 at midnight (00:00:00)
//! - **Performance timing**: Ideal for measuring elapsed time for operations
//! - **System dependent**: Precision varies by platform and Windows version
//! - **Time vs Timer**: `Time` returns Date type for clock time, `Timer` returns Single for measurements
//! - **Midnight handling**: When crossing midnight, current - start will be negative
//! - **Maximum value**: Approximately 86,400 (24 hours × 3600 seconds/hour)
//!
//! ### Timer vs Related Functions
//! - `Timer` - Returns seconds since midnight as Single (high precision)
//! - `Time` - Returns current time as Date (1 second precision)
//! - `Now` - Returns current date and time as Date
//! - `GetTickCount` - Windows API function for milliseconds since system start
//!
//! ### Precision Notes
//! - Windows 95/98/ME: ~55 milliseconds
//! - Windows NT/2000/XP and later: ~10-15 milliseconds
//! - Not suitable for microsecond precision measurements
//! - Use QueryPerformanceCounter API for higher precision
//!
//! ### Midnight Rollover Handling
//! ```vb6
//! Function SafeElapsedTime(startTime As Single) As Single
//!     Dim elapsed As Single
//!     elapsed = Timer - startTime
//!     
//!     ' Handle midnight rollover
//!     If elapsed < 0 Then
//!         elapsed = elapsed + 86400   ' Add 24 hours worth of seconds
//!     End If
//!     
//!     SafeElapsedTime = elapsed
//! End Function
//! ```
//!
//! ## Typical Uses
//! 1. **Performance Measurement**: Time how long operations take
//! 2. **Timeout Implementation**: Check if time limit exceeded
//! 3. **Animation Timing**: Control animation frame timing
//! 4. **Delay Implementation**: Create precise delays
//! 5. **Benchmark Testing**: Compare performance of different approaches
//! 6. **Rate Limiting**: Control operation frequency
//! 7. **Elapsed Time Display**: Show how long process has been running
//! 8. **Time-based Triggers**: Execute code after specific duration
//!
//! ## Basic Examples
//!
//! ### Example 1: Measure Operation Time
//! ```vb6
//! Sub MeasureOperation()
//!     Dim startTime As Single
//!     Dim elapsed As Single
//!     
//!     startTime = Timer
//!     
//!     ' Perform operation
//!     Call LongRunningOperation
//!     
//!     elapsed = Timer - startTime
//!     MsgBox "Operation took " & Format$(elapsed, "0.000") & " seconds"
//! End Sub
//! ```
//!
//! ### Example 2: Simple Timeout
//! ```vb6
//! Sub WaitWithTimeout(seconds As Single)
//!     Dim startTime As Single
//!     startTime = Timer
//!     
//!     Do While Timer - startTime < seconds
//!         DoEvents  ' Allow other processing
//!     Loop
//! End Sub
//! ```
//!
//! ### Example 3: Check If Timeout Exceeded
//! ```vb6
//! Function IsTimeout(startTime As Single, timeoutSeconds As Single) As Boolean
//!     IsTimeout = (Timer - startTime >= timeoutSeconds)
//! End Function
//! ```
//!
//! ### Example 4: Benchmark Comparison
//! ```vb6
//! Sub CompareMethods()
//!     Dim time1 As Single, time2 As Single
//!     Dim elapsed1 As Single, elapsed2 As Single
//!     
//!     ' Test Method 1
//!     time1 = Timer
//!     Call Method1
//!     elapsed1 = Timer - time1
//!     
//!     ' Test Method 2
//!     time2 = Timer
//!     Call Method2
//!     elapsed2 = Timer - time2
//!     
//!     Debug.Print "Method1: " & elapsed1 & "s, Method2: " & elapsed2 & "s"
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Basic Elapsed Time
//! ```vb6
//! Function GetElapsedTime(startTime As Single) As Single
//!     GetElapsedTime = Timer - startTime
//! End Function
//! ```
//!
//! ### Pattern 2: Elapsed Time with Midnight Handling
//! ```vb6
//! Function GetElapsedTimeSafe(startTime As Single) As Single
//!     Dim elapsed As Single
//!     elapsed = Timer - startTime
//!     If elapsed < 0 Then elapsed = elapsed + 86400
//!     GetElapsedTimeSafe = elapsed
//! End Function
//! ```
//!
//! ### Pattern 3: Wait with DoEvents
//! ```vb6
//! Sub WaitSeconds(seconds As Single)
//!     Dim endTime As Single
//!     endTime = Timer + seconds
//!     
//!     Do While Timer < endTime
//!         DoEvents
//!     Loop
//! End Sub
//! ```
//!
//! ### Pattern 4: Timeout Loop
//! ```vb6
//! Function WaitForCondition(timeoutSeconds As Single) As Boolean
//!     Dim startTime As Single
//!     startTime = Timer
//!     
//!     Do While Not ConditionMet()
//!         If Timer - startTime > timeoutSeconds Then
//!             WaitForCondition = False
//!             Exit Function
//!         End If
//!         DoEvents
//!     Loop
//!     
//!     WaitForCondition = True
//! End Function
//! ```
//!
//! ### Pattern 5: Frame Rate Control
//! ```vb6
//! Sub ControlFrameRate(targetFPS As Single)
//!     Static lastFrameTime As Single
//!     Dim targetDelay As Single
//!     Dim elapsed As Single
//!     
//!     targetDelay = 1 / targetFPS
//!     elapsed = Timer - lastFrameTime
//!     
//!     If elapsed < targetDelay Then
//!         ' Wait for remaining time
//!         Do While Timer - lastFrameTime < targetDelay
//!             DoEvents
//!         Loop
//!     End If
//!     
//!     lastFrameTime = Timer
//! End Sub
//! ```
//!
//! ### Pattern 6: Rate Limiter
//! ```vb6
//! Function CanExecute(minimumInterval As Single) As Boolean
//!     Static lastExecuteTime As Single
//!     
//!     If Timer - lastExecuteTime >= minimumInterval Then
//!         lastExecuteTime = Timer
//!         CanExecute = True
//!     Else
//!         CanExecute = False
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: Performance Counter
//! ```vb6
//! Sub StartTimer()
//!     Static timerStart As Single
//!     timerStart = Timer
//! End Sub
//!
//! Function GetTimerElapsed() As Single
//!     Static timerStart As Single
//!     GetTimerElapsed = Timer - timerStart
//! End Function
//! ```
//!
//! ### Pattern 8: Format Elapsed Time
//! ```vb6
//! Function FormatElapsed(seconds As Single) As String
//!     Dim hrs As Long, mins As Long, secs As Long
//!     
//!     hrs = Int(seconds / 3600)
//!     mins = Int((seconds - hrs * 3600) / 60)
//!     secs = Int(seconds - hrs * 3600 - mins * 60)
//!     
//!     FormatElapsed = Format$(hrs, "00") & ":" & _
//!                    Format$(mins, "00") & ":" & _
//!                    Format$(secs, "00")
//! End Function
//! ```
//!
//! ### Pattern 9: Delay with Accuracy Check
//! ```vb6
//! Sub AccurateDelay(milliseconds As Long)
//!     Dim targetTime As Single
//!     targetTime = Timer + (milliseconds / 1000)
//!     
//!     Do While Timer < targetTime
//!         ' Busy wait for precision
//!     Loop
//! End Sub
//! ```
//!
//! ### Pattern 10: Periodic Execution
//! ```vb6
//! Sub ExecutePeriodically(intervalSeconds As Single)
//!     Static lastRun As Single
//!     
//!     If Timer - lastRun >= intervalSeconds Or lastRun = 0 Then
//!         ' Execute periodic task
//!         PerformTask
//!         lastRun = Timer
//!     End If
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Performance Profiler Class
//! ```vb6
//! ' Class: PerformanceProfiler
//! ' Profiles code execution times with statistics
//! Option Explicit
//!
//! Private m_StartTime As Single
//! Private m_Measurements As Collection
//! Private m_IsRunning As Boolean
//!
//! Private Sub Class_Initialize()
//!     Set m_Measurements = New Collection
//! End Sub
//!
//! Public Sub Start()
//!     m_StartTime = Timer
//!     m_IsRunning = True
//! End Sub
//!
//! Public Function Stop() As Single
//!     Dim elapsed As Single
//!     
//!     If Not m_IsRunning Then
//!         Err.Raise 5, , "Profiler not started"
//!     End If
//!     
//!     elapsed = Timer - m_StartTime
//!     If elapsed < 0 Then elapsed = elapsed + 86400  ' Handle midnight
//!     
//!     m_Measurements.Add elapsed
//!     m_IsRunning = False
//!     Stop = elapsed
//! End Function
//!
//! Public Function GetAverageTime() As Single
//!     Dim total As Single
//!     Dim measurement As Variant
//!     
//!     If m_Measurements.Count = 0 Then
//!         GetAverageTime = 0
//!         Exit Function
//!     End If
//!     
//!     total = 0
//!     For Each measurement In m_Measurements
//!         total = total + measurement
//!     Next measurement
//!     
//!     GetAverageTime = total / m_Measurements.Count
//! End Function
//!
//! Public Function GetMinTime() As Single
//!     Dim minTime As Single
//!     Dim measurement As Variant
//!     
//!     If m_Measurements.Count = 0 Then
//!         GetMinTime = 0
//!         Exit Function
//!     End If
//!     
//!     minTime = 999999
//!     For Each measurement In m_Measurements
//!         If measurement < minTime Then minTime = measurement
//!     Next measurement
//!     
//!     GetMinTime = minTime
//! End Function
//!
//! Public Function GetMaxTime() As Single
//!     Dim maxTime As Single
//!     Dim measurement As Variant
//!     
//!     If m_Measurements.Count = 0 Then
//!         GetMaxTime = 0
//!         Exit Function
//!     End If
//!     
//!     maxTime = 0
//!     For Each measurement In m_Measurements
//!         If measurement > maxTime Then maxTime = measurement
//!     Next measurement
//!     
//!     GetMaxTime = maxTime
//! End Function
//!
//! Public Sub Reset()
//!     Set m_Measurements = New Collection
//!     m_IsRunning = False
//! End Sub
//!
//! Public Property Get MeasurementCount() As Long
//!     MeasurementCount = m_Measurements.Count
//! End Property
//! ```
//!
//! ### Example 2: Timeout Manager Module
//! ```vb6
//! ' Module: TimeoutManager
//! ' Manages multiple concurrent timeouts
//! Option Explicit
//!
//! Private Type TimeoutEntry
//!     Name As String
//!     StartTime As Single
//!     TimeoutSeconds As Single
//!     Active As Boolean
//! End Type
//!
//! Private m_Timeouts() As TimeoutEntry
//! Private m_TimeoutCount As Long
//!
//! Public Sub StartTimeout(name As String, timeoutSeconds As Single)
//!     Dim i As Long
//!     
//!     ' Check if already exists
//!     For i = 0 To m_TimeoutCount - 1
//!         If m_Timeouts(i).Name = name Then
//!             m_Timeouts(i).StartTime = Timer
//!             m_Timeouts(i).TimeoutSeconds = timeoutSeconds
//!             m_Timeouts(i).Active = True
//!             Exit Sub
//!         End If
//!     Next i
//!     
//!     ' Add new timeout
//!     ReDim Preserve m_Timeouts(m_TimeoutCount)
//!     m_Timeouts(m_TimeoutCount).Name = name
//!     m_Timeouts(m_TimeoutCount).StartTime = Timer
//!     m_Timeouts(m_TimeoutCount).TimeoutSeconds = timeoutSeconds
//!     m_Timeouts(m_TimeoutCount).Active = True
//!     m_TimeoutCount = m_TimeoutCount + 1
//! End Sub
//!
//! Public Function IsTimedOut(name As String) As Boolean
//!     Dim i As Long
//!     Dim elapsed As Single
//!     
//!     For i = 0 To m_TimeoutCount - 1
//!         If m_Timeouts(i).Name = name And m_Timeouts(i).Active Then
//!             elapsed = Timer - m_Timeouts(i).StartTime
//!             If elapsed < 0 Then elapsed = elapsed + 86400
//!             
//!             IsTimedOut = (elapsed >= m_Timeouts(i).TimeoutSeconds)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsTimedOut = False
//! End Function
//!
//! Public Sub CancelTimeout(name As String)
//!     Dim i As Long
//!     
//!     For i = 0 To m_TimeoutCount - 1
//!         If m_Timeouts(i).Name = name Then
//!             m_Timeouts(i).Active = False
//!             Exit Sub
//!         End If
//!     Next i
//! End Sub
//!
//! Public Function GetTimeRemaining(name As String) As Single
//!     Dim i As Long
//!     Dim elapsed As Single
//!     
//!     For i = 0 To m_TimeoutCount - 1
//!         If m_Timeouts(i).Name = name And m_Timeouts(i).Active Then
//!             elapsed = Timer - m_Timeouts(i).StartTime
//!             If elapsed < 0 Then elapsed = elapsed + 86400
//!             
//!             GetTimeRemaining = m_Timeouts(i).TimeoutSeconds - elapsed
//!             If GetTimeRemaining < 0 Then GetTimeRemaining = 0
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetTimeRemaining = 0
//! End Function
//! ```
//!
//! ### Example 3: Animation Timer Class
//! ```vb6
//! ' Class: AnimationTimer
//! ' Controls animation frame timing and FPS
//! Option Explicit
//!
//! Private m_TargetFPS As Single
//! Private m_LastFrameTime As Single
//! Private m_FrameCount As Long
//! Private m_FPSCalculationTime As Single
//! Private m_CurrentFPS As Single
//!
//! Public Sub Initialize(targetFPS As Single)
//!     m_TargetFPS = targetFPS
//!     m_LastFrameTime = Timer
//!     m_FPSCalculationTime = Timer
//!     m_FrameCount = 0
//!     m_CurrentFPS = 0
//! End Sub
//!
//! Public Function ShouldRenderFrame() As Boolean
//!     Dim currentTime As Single
//!     Dim targetDelay As Single
//!     Dim elapsed As Single
//!     
//!     currentTime = Timer
//!     targetDelay = 1 / m_TargetFPS
//!     
//!     elapsed = currentTime - m_LastFrameTime
//!     If elapsed < 0 Then elapsed = elapsed + 86400
//!     
//!     If elapsed >= targetDelay Then
//!         m_LastFrameTime = currentTime
//!         m_FrameCount = m_FrameCount + 1
//!         
//!         ' Update FPS calculation every second
//!         If currentTime - m_FPSCalculationTime >= 1 Then
//!             m_CurrentFPS = m_FrameCount / (currentTime - m_FPSCalculationTime)
//!             m_FrameCount = 0
//!             m_FPSCalculationTime = currentTime
//!         End If
//!         
//!         ShouldRenderFrame = True
//!     Else
//!         ShouldRenderFrame = False
//!     End If
//! End Function
//!
//! Public Property Get CurrentFPS() As Single
//!     CurrentFPS = m_CurrentFPS
//! End Property
//!
//! Public Property Get TargetFPS() As Single
//!     TargetFPS = m_TargetFPS
//! End Property
//!
//! Public Property Let TargetFPS(value As Single)
//!     If value > 0 Then m_TargetFPS = value
//! End Property
//! ```
//!
//! ### Example 4: Benchmark Suite Module
//! ```vb6
//! ' Module: BenchmarkSuite
//! ' Runs and compares multiple benchmark tests
//! Option Explicit
//!
//! Private Type BenchmarkResult
//!     Name As String
//!     Iterations As Long
//!     TotalTime As Single
//!     AverageTime As Single
//!     MinTime As Single
//!     MaxTime As Single
//! End Type
//!
//! Public Function RunBenchmark(testName As String, iterations As Long) As BenchmarkResult
//!     Dim result As BenchmarkResult
//!     Dim i As Long
//!     Dim startTime As Single
//!     Dim elapsed As Single
//!     Dim minTime As Single
//!     Dim maxTime As Single
//!     Dim totalTime As Single
//!     
//!     result.Name = testName
//!     result.Iterations = iterations
//!     minTime = 999999
//!     maxTime = 0
//!     totalTime = 0
//!     
//!     For i = 1 To iterations
//!         startTime = Timer
//!         
//!         ' Execute test
//!         Call ExecuteBenchmarkTest(testName)
//!         
//!         elapsed = Timer - startTime
//!         If elapsed < 0 Then elapsed = elapsed + 86400
//!         
//!         totalTime = totalTime + elapsed
//!         If elapsed < minTime Then minTime = elapsed
//!         If elapsed > maxTime Then maxTime = elapsed
//!     Next i
//!     
//!     result.TotalTime = totalTime
//!     result.AverageTime = totalTime / iterations
//!     result.MinTime = minTime
//!     result.MaxTime = maxTime
//!     
//!     RunBenchmark = result
//! End Function
//!
//! Public Function FormatBenchmarkResult(result As BenchmarkResult) As String
//!     Dim output As String
//!     
//!     output = "Benchmark: " & result.Name & vbCrLf
//!     output = output & "Iterations: " & result.Iterations & vbCrLf
//!     output = output & "Total Time: " & Format$(result.TotalTime, "0.0000") & "s" & vbCrLf
//!     output = output & "Average: " & Format$(result.AverageTime * 1000, "0.000") & "ms" & vbCrLf
//!     output = output & "Min: " & Format$(result.MinTime * 1000, "0.000") & "ms" & vbCrLf
//!     output = output & "Max: " & Format$(result.MaxTime * 1000, "0.000") & "ms" & vbCrLf
//!     
//!     FormatBenchmarkResult = output
//! End Function
//!
//! Private Sub ExecuteBenchmarkTest(testName As String)
//!     ' Placeholder - implement actual test logic
//! End Sub
//! ```
//!
//! ## Error Handling
//! The `Timer` function typically does not raise errors:
//!
//! - **No parameters**: Cannot have parameter errors
//! - **Always succeeds**: Returns current timer value
//! - **Midnight rollover**: Not an error, but requires handling in elapsed time calculations
//! - **Negative elapsed time**: Indicates midnight rollover, add 86400 to correct
//!
//! ## Performance Notes
//! - Very fast operation - direct system call
//! - Minimal CPU overhead
//! - Precision varies by Windows version (10-55ms)
//! - Not suitable for microsecond-level timing
//! - For busy-wait loops, consider CPU usage impact
//! - Use DoEvents in wait loops to prevent UI freeze
//!
//! ## Best Practices
//! 1. **Handle midnight rollover** when measuring elapsed time
//! 2. **Use DoEvents** in timing loops to prevent application freeze
//! 3. **Store as Single** type for consistency and precision
//! 4. **Avoid long-running measurements** crossing midnight
//! 5. **Use for relative timing** not absolute time-of-day
//! 6. **Consider GetTickCount API** for measurements exceeding 24 hours
//! 7. **Test across midnight** boundary for production code
//! 8. **Use Time function** for actual clock time display
//! 9. **Format appropriately** when displaying to users (ms, seconds, minutes)
//! 10. **Profile performance** in release builds, not debug mode
//!
//! ## Comparison Table
//!
//! | Function | Returns | Precision | Range | Purpose |
//! |----------|---------|-----------|-------|---------|
//! | `Timer` | Single | 10-55ms | 0-86400 | Performance timing |
//! | `Time` | Date | 1 second | Full time | Clock time |
//! | `Now` | Date | 1 second | Full date/time | Current date/time |
//! | `GetTickCount` | Long | 10-16ms | 0-49.7 days | Uptime timing |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and VBScript
//! - Precision varies by Windows version
//! - Windows 95/98/ME: ~55ms resolution
//! - Windows NT/2000/XP+: ~10-15ms resolution
//! - Returns fractional seconds (not milliseconds)
//! - Resets at midnight (00:00:00)
//! - Based on system timer tick
//!
//! ## Limitations
//! - Resets to 0 at midnight (requires handling)
//! - Not suitable for measurements exceeding 24 hours
//! - Limited precision (10-55ms, not microseconds)
//! - No timezone awareness
//! - Cannot measure absolute time, only relative
//! - Precision varies between systems
//! - Not monotonic (can jump backwards at midnight)
//! - No built-in high-resolution timer (use QueryPerformanceCounter API)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_timer_basic() {
        let source = r#"
Sub Test()
    elapsed = Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_with_parentheses() {
        let source = r#"
Sub Test()
    elapsed = Timer()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_variable_assignment() {
        let source = r#"
Sub Test()
    Dim startTime As Single
    startTime = Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_elapsed_time() {
        let source = r#"
Sub Test()
    elapsed = Timer - startTime
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_function_return() {
        let source = r#"
Function GetStartTime() As Single
    GetStartTime = Timer
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Elapsed: " & (Timer - startTime)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_if_statement() {
        let source = r#"
Sub Test()
    If Timer - startTime > 10 Then
        Timeout
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_comparison() {
        let source = r#"
Sub Test()
    If Timer < endTime Then
        Continue
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_do_while() {
        let source = r#"
Sub Test()
    Do While Timer - startTime < 5
        DoEvents
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_do_until() {
        let source = r#"
Sub Test()
    Do Until Timer - startTime >= timeout
        Process
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_while_wend() {
        let source = r#"
Sub Test()
    While Timer < targetTime
        Wait
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_function_argument() {
        let source = r#"
Sub Test()
    Call LogTime(Timer)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_property_assignment() {
        let source = r#"
Sub Test()
    obj.StartTime = Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_with_statement() {
        let source = r#"
Sub Test()
    With profiler
        .Start = Timer
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_array_assignment() {
        let source = r#"
Sub Test()
    timestamps(i) = Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_print_statement() {
        let source = r#"
Sub Test()
    Print #1, Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_format() {
        let source = r#"
Sub Test()
    formatted = Format$(Timer - startTime, "0.000")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_addition() {
        let source = r#"
Sub Test()
    targetTime = Timer + 10
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf Timer - startTime > 5 Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_select_case() {
        let source = r#"
Sub Test()
    Select Case Int(Timer - startTime)
        Case 0 To 5
            Fast
        Case 6 To 10
            Medium
        Case Else
            Slow
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_iif() {
        let source = r#"
Sub Test()
    status = IIf(Timer - startTime > timeout, "Timeout", "OK")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_midnight_handling() {
        let source = r#"
Sub Test()
    elapsed = Timer - startTime
    If elapsed < 0 Then elapsed = elapsed + 86400
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_class_usage() {
        let source = r#"
Sub Test()
    Set profiler = New PerformanceProfiler
    profiler.Start Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_csng() {
        let source = r#"
Sub Test()
    timerValue = CSng(Timer)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_parentheses_expression() {
        let source = r#"
Sub Test()
    result = (Timer - startTime) * 1000
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_concatenation() {
        let source = r#"
Sub Test()
    logEntry = "Time: " & Timer
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn test_timer_multiple_usage() {
        let source = r#"
Sub Test()
    start1 = Timer
    DoSomething
    end1 = Timer
    elapsed = end1 - start1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Timer"));
    }
}
