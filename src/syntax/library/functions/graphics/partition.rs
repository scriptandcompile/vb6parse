//! # Partition Function
//!
//! Returns a String indicating where a number occurs within a calculated series of ranges.
//!
//! ## Syntax
//!
//! ```vb
//! Partition(number, start, stop, interval)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Whole number that you want to locate within one of the ranges.
//! - `start` - Required. Whole number that is the start of the overall range of numbers. Cannot be less than 0.
//! - `stop` - Required. Whole number that is the end of the overall range of numbers. Cannot be equal to or less than `start`.
//! - `interval` - Required. Whole number that indicates the size of each range from `start` to `stop`. Cannot be less than 1.
//!
//! ## Return Value
//!
//! Returns a `String` describing the range in which `number` falls. The format is:
//! - `"lowerbound:upperbound"` for ranges within the series
//! - `" :lowerbound-1"` for values less than `start`
//! - `"upperbound+1: "` for values greater than `stop`
//!
//! ## Remarks
//!
//! The `Partition` function divides a range of numbers into smaller intervals and returns a string
//! describing which interval contains a given number. This is particularly useful for creating
//! frequency distributions, histograms, and grouping data into bins.
//!
//! The function creates ranges starting at `start` and ending at `stop`, with each range having
//! a width of `interval`. The returned string always has the same width for both the lower and
//! upper boundaries, padded with leading spaces as needed. This ensures consistent formatting
//! when building reports or tables.
//!
//! For example, if `start` is 0, `stop` is 100, and `interval` is 10, the function creates
//! ranges: 0:9, 10:19, 20:29, etc., plus special ranges for values below 0 (":–1") and
//! above 100 ("101:").
//!
//! The width of each number in the returned string is calculated based on the number of digits
//! in `stop` plus 1. This ensures all ranges align properly in columnar displays.
//!
//! ## Typical Uses
//!
//! 1. **Frequency Distribution**: Creating frequency tables for statistical analysis
//! 2. **Histogram Generation**: Grouping data into bins for histogram charts
//! 3. **Age Group Analysis**: Categorizing people into age brackets (0-9, 10-19, etc.)
//! 4. **Sales Range Reports**: Grouping sales figures into ranges for analysis
//! 5. **Performance Banding**: Categorizing test scores or metrics into performance bands
//! 6. **Data Binning**: Organizing continuous data into discrete categories
//! 7. **Time Period Grouping**: Grouping timestamps into hour, day, or week ranges
//! 8. **Price Range Analysis**: Categorizing products by price ranges
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Partition
//! ```vb
//! Dim range As String
//! range = Partition(15, 0, 100, 10)    ' Returns " 10: 19"
//! range = Partition(5, 0, 100, 10)     ' Returns "  0:  9"
//! range = Partition(95, 0, 100, 10)    ' Returns " 90: 99"
//! ```
//!
//! ### Example 2: Frequency Distribution
//! ```vb
//! ' Count how many values fall in each range
//! Dim values(100) As Integer
//! Dim frequency As Collection
//! Dim i As Integer
//! Dim range As String
//!
//! Set frequency = New Collection
//!
//! ' Populate with sample data
//! For i = 0 To 100
//!     values(i) = Int(Rnd * 100)
//! Next i
//!
//! ' Count frequencies
//! For i = 0 To 100
//!     range = Partition(values(i), 0, 99, 10)
//!     ' Increment count for this range
//! Next i
//! ```
//!
//! ### Example 3: Age Grouping
//! ```vb
//! Function GetAgeGroup(age As Integer) As String
//!     ' Group ages into decades
//!     GetAgeGroup = Partition(age, 0, 100, 10)
//! End Function
//!
//! ' Usage
//! Debug.Print GetAgeGroup(25)    ' Returns " 20: 29"
//! Debug.Print GetAgeGroup(5)     ' Returns "  0:  9"
//! ```
//!
//! ### Example 4: Out of Range Values
//! ```vb
//! Dim range As String
//! range = Partition(-5, 0, 100, 10)    ' Returns "   : -1" (below start)
//! range = Partition(150, 0, 100, 10)   ' Returns "101:   " (above stop)
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `BuildFrequencyTable`
//! ```vb
//! Function BuildFrequencyTable(values() As Integer, start As Long, _
//!                              stop As Long, interval As Long) As Collection
//!     Dim i As Integer
//!     Dim range As String
//!     Dim count As Long
//!     Dim freq As Collection
//!     
//!     Set freq = New Collection
//!     
//!     ' Initialize all ranges to 0
//!     For i = start To stop Step interval
//!         range = Partition(i, start, stop, interval)
//!         On Error Resume Next
//!         freq.Add 0, range
//!         On Error GoTo 0
//!     Next i
//!     
//!     ' Count occurrences
//!     For i = LBound(values) To UBound(values)
//!         range = Partition(values(i), start, stop, interval)
//!         On Error Resume Next
//!         count = freq(range)
//!         freq.Remove range
//!         freq.Add count + 1, range
//!         On Error GoTo 0
//!     Next i
//!     
//!     Set BuildFrequencyTable = freq
//! End Function
//! ```
//!
//! ### Pattern 2: `GenerateHistogram`
//! ```vb
//! Sub GenerateHistogram(values() As Integer, start As Long, _
//!                       stop As Long, interval As Long)
//!     Dim i As Integer
//!     Dim range As String
//!     Dim counts As Object
//!     Dim currentRange As String
//!     
//!     Set counts = CreateObject("Scripting.Dictionary")
//!     
//!     ' Count frequencies
//!     For i = LBound(values) To UBound(values)
//!         range = Partition(values(i), start, stop, interval)
//!         If Not counts.Exists(range) Then
//!             counts.Add range, 0
//!         End If
//!         counts(range) = counts(range) + 1
//!     Next i
//!     
//!     ' Display histogram
//!     Debug.Print "Range", "Count", "Chart"
//!     Debug.Print String(50, "-")
//!     For Each currentRange In counts.Keys
//!         Debug.Print currentRange, counts(currentRange), _
//!                     String(counts(currentRange), "*")
//!     Next currentRange
//! End Sub
//! ```
//!
//! ### Pattern 3: `ClassifyValue`
//! ```vb
//! Function ClassifyValue(value As Long, start As Long, _
//!                        stop As Long, interval As Long) As String
//!     Dim range As String
//!     range = Partition(value, start, stop, interval)
//!     
//!     If InStr(range, ":") = 1 Then
//!         ClassifyValue = "Below Range"
//!     ElseIf Right(range, 1) = ":" Then
//!         ClassifyValue = "Above Range"
//!     Else
//!         ClassifyValue = "In Range: " & Trim(range)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `GetRangeBoundaries`
//! ```vb
//! Sub GetRangeBoundaries(partitionStr As String, _
//!                        ByRef lower As Long, ByRef upper As Long)
//!     Dim parts() As String
//!     parts = Split(partitionStr, ":")
//!     
//!     If UBound(parts) = 1 Then
//!         On Error Resume Next
//!         lower = CLng(Trim(parts(0)))
//!         upper = CLng(Trim(parts(1)))
//!         On Error GoTo 0
//!     End If
//! End Sub
//! ' Usage: GetRangeBoundaries(" 20: 29", lower, upper)  ' lower=20, upper=29
//! ```
//!
//! ### Pattern 5: `AnalyzeDataDistribution`
//! ```vb
//! Function AnalyzeDataDistribution(data() As Variant) As String
//!     Dim minVal As Long, maxVal As Long
//!     Dim i As Long
//!     Dim report As String
//!     Dim range As String
//!     Dim counts As Object
//!     
//!     ' Find min and max
//!     minVal = data(LBound(data))
//!     maxVal = data(LBound(data))
//!     For i = LBound(data) + 1 To UBound(data)
//!         If data(i) < minVal Then minVal = data(i)
//!         If data(i) > maxVal Then maxVal = data(i)
//!     Next i
//!     
//!     ' Build frequency table
//!     Set counts = CreateObject("Scripting.Dictionary")
//!     For i = LBound(data) To UBound(data)
//!         range = Partition(data(i), minVal, maxVal, (maxVal - minVal) \ 10)
//!         If Not counts.Exists(range) Then counts.Add range, 0
//!         counts(range) = counts(range) + 1
//!     Next i
//!     
//!     ' Build report
//!     report = "Data Distribution:" & vbCrLf
//!     For Each range In counts.Keys
//!         report = report & range & ": " & counts(range) & vbCrLf
//!     Next range
//!     
//!     AnalyzeDataDistribution = report
//! End Function
//! ```
//!
//! ### Pattern 6: `ValidatePartitionParameters`
//! ```vb
//! Function ValidatePartitionParameters(start As Long, stop As Long, _
//!                                      interval As Long) As Boolean
//!     ValidatePartitionParameters = False
//!     
//!     If start < 0 Then
//!         MsgBox "Start must be >= 0"
//!         Exit Function
//!     End If
//!     
//!     If stop <= start Then
//!         MsgBox "Stop must be > start"
//!         Exit Function
//!     End If
//!     
//!     If interval < 1 Then
//!         MsgBox "Interval must be >= 1"
//!         Exit Function
//!     End If
//!     
//!     ValidatePartitionParameters = True
//! End Function
//! ```
//!
//! ### Pattern 7: `CreateRangeLabels`
//! ```vb
//! Function CreateRangeLabels(start As Long, stop As Long, _
//!                            interval As Long) As String()
//!     Dim labels() As String
//!     Dim count As Long
//!     Dim i As Long
//!     Dim index As Long
//!     
//!     count = (stop - start) \ interval + 3  ' +3 for below, above, and safety
//!     ReDim labels(0 To count)
//!     
//!     index = 0
//!     For i = start To stop Step interval
//!         labels(index) = Partition(i, start, stop, interval)
//!         index = index + 1
//!     Next i
//!     
//!     ReDim Preserve labels(0 To index - 1)
//!     CreateRangeLabels = labels
//! End Function
//! ```
//!
//! ### Pattern 8: `CountInRange`
//! ```vb
//! Function CountInRange(values() As Variant, targetRange As String, _
//!                       start As Long, stop As Long, interval As Long) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     count = 0
//!     For i = LBound(values) To UBound(values)
//!         If Partition(values(i), start, stop, interval) = targetRange Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountInRange = count
//! End Function
//! ```
//!
//! ### Pattern 9: `GetRangeMidpoint`
//! ```vb
//! Function GetRangeMidpoint(partitionStr As String) As Double
//!     Dim lower As Long, upper As Long
//!     Dim parts() As String
//!     
//!     parts = Split(partitionStr, ":")
//!     If UBound(parts) = 1 Then
//!         On Error Resume Next
//!         lower = CLng(Trim(parts(0)))
//!         upper = CLng(Trim(parts(1)))
//!         If Err.Number = 0 Then
//!             GetRangeMidpoint = (lower + upper) / 2
//!         End If
//!         On Error GoTo 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `GroupByRange`
//! ```vb
//! Function GroupByRange(values() As Variant, start As Long, _
//!                       stop As Long, interval As Long) As Object
//!     Dim groups As Object
//!     Dim i As Long
//!     Dim range As String
//!     Dim groupItems As Collection
//!     
//!     Set groups = CreateObject("Scripting.Dictionary")
//!     
//!     For i = LBound(values) To UBound(values)
//!         range = Partition(values(i), start, stop, interval)
//!         
//!         If Not groups.Exists(range) Then
//!             Set groupItems = New Collection
//!             groups.Add range, groupItems
//!         End If
//!         
//!         groups(range).Add values(i)
//!     Next i
//!     
//!     Set GroupByRange = groups
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Statistical Analysis Tool
//! ```vb
//! ' Comprehensive statistical analysis using Partition
//! Class StatisticalAnalyzer
//!     Private m_data() As Double
//!     Private m_binCount As Long
//!     Private m_minValue As Double
//!     Private m_maxValue As Double
//!     
//!     Public Sub LoadData(data() As Double)
//!         Dim i As Long
//!         ReDim m_data(LBound(data) To UBound(data))
//!         
//!         m_minValue = data(LBound(data))
//!         m_maxValue = data(LBound(data))
//!         
//!         For i = LBound(data) To UBound(data)
//!             m_data(i) = data(i)
//!             If data(i) < m_minValue Then m_minValue = data(i)
//!             If data(i) > m_maxValue Then m_maxValue = data(i)
//!         Next i
//!     End Sub
//!     
//!     Public Property Let BinCount(value As Long)
//!         If value > 0 Then m_binCount = value
//!     End Property
//!     
//!     Public Function GetFrequencyDistribution() As Object
//!         Dim freq As Object
//!         Dim i As Long
//!         Dim interval As Long
//!         Dim range As String
//!         Dim start As Long, stop As Long
//!         
//!         Set freq = CreateObject("Scripting.Dictionary")
//!         
//!         start = Int(m_minValue)
//!         stop = Int(m_maxValue)
//!         interval = (stop - start) \ m_binCount
//!         If interval < 1 Then interval = 1
//!         
//!         For i = LBound(m_data) To UBound(m_data)
//!             range = Partition(Int(m_data(i)), start, stop, interval)
//!             If Not freq.Exists(range) Then
//!                 freq.Add range, 0
//!             End If
//!             freq(range) = freq(range) + 1
//!         Next i
//!         
//!         Set GetFrequencyDistribution = freq
//!     End Function
//!     
//!     Public Function GetHistogramData() As Object
//!         Dim histogram As Object
//!         Dim freq As Object
//!         Dim range As Variant
//!         Dim item As Object
//!         
//!         Set freq = GetFrequencyDistribution()
//!         Set histogram = CreateObject("Scripting.Dictionary")
//!         
//!         For Each range In freq.Keys
//!             Set item = CreateObject("Scripting.Dictionary")
//!             item.Add "Range", range
//!             item.Add "Count", freq(range)
//!             item.Add "Percentage", (freq(range) / UBound(m_data)) * 100
//!             histogram.Add range, item
//!         Next range
//!         
//!         Set GetHistogramData = histogram
//!     End Function
//!     
//!     Public Function GenerateReport() As String
//!         Dim report As String
//!         Dim histogram As Object
//!         Dim range As Variant
//!         Dim maxCount As Long
//!         Dim barWidth As Long
//!         
//!         Set histogram = GetHistogramData()
//!         
//!         ' Find max count for scaling
//!         maxCount = 0
//!         For Each range In histogram.Keys
//!             If histogram(range)("Count") > maxCount Then
//!                 maxCount = histogram(range)("Count")
//!             End If
//!         Next range
//!         
//!         report = "Frequency Distribution Report" & vbCrLf
//!         report = report & String(60, "=") & vbCrLf
//!         report = report & "Data Points: " & UBound(m_data) + 1 & vbCrLf
//!         report = report & "Min Value: " & m_minValue & vbCrLf
//!         report = report & "Max Value: " & m_maxValue & vbCrLf
//!         report = report & String(60, "-") & vbCrLf
//!         report = report & "Range          Count    Pct    Chart" & vbCrLf
//!         report = report & String(60, "-") & vbCrLf
//!         
//!         For Each range In histogram.Keys
//!             barWidth = Int((histogram(range)("Count") / maxCount) * 30)
//!             report = report & range & "  " & _
//!                      Format(histogram(range)("Count"), "0000") & "  " & _
//!                      Format(histogram(range)("Percentage"), "00.0") & "%  " & _
//!                      String(barWidth, "#") & vbCrLf
//!         Next range
//!         
//!         GenerateReport = report
//!     End Function
//!     
//!     Public Function GetModeRange() As String
//!         ' Find the range with the highest frequency
//!         Dim freq As Object
//!         Dim range As Variant
//!         Dim maxCount As Long
//!         Dim modeRange As String
//!         
//!         Set freq = GetFrequencyDistribution()
//!         maxCount = 0
//!         
//!         For Each range In freq.Keys
//!             If freq(range) > maxCount Then
//!                 maxCount = freq(range)
//!                 modeRange = range
//!             End If
//!         Next range
//!         
//!         GetModeRange = modeRange
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Sales Performance Analyzer
//! ```vb
//! ' Analyze sales performance by grouping into performance bands
//! Module SalesAnalyzer
//!     Private Type SalesRecord
//!         SalesPerson As String
//!         Amount As Currency
//!         Date As Date
//!     End Type
//!     
//!     Public Function AnalyzeSalesPerformance(sales() As SalesRecord) As String
//!         Dim i As Long
//!         Dim minSale As Currency, maxSale As Currency
//!         Dim interval As Currency
//!         Dim performanceBands As Object
//!         Dim range As String
//!         Dim report As String
//!         Dim totalSales As Long
//!         
//!         ' Find range
//!         minSale = sales(LBound(sales)).Amount
//!         maxSale = sales(LBound(sales)).Amount
//!         For i = LBound(sales) To UBound(sales)
//!             If sales(i).Amount < minSale Then minSale = sales(i).Amount
//!             If sales(i).Amount > maxSale Then maxSale = sales(i).Amount
//!         Next i
//!         
//!         ' Create 5 performance bands
//!         interval = Int((maxSale - minSale) / 5)
//!         If interval < 1 Then interval = 1
//!         
//!         Set performanceBands = CreateObject("Scripting.Dictionary")
//!         
//!         ' Categorize sales
//!         totalSales = UBound(sales) - LBound(sales) + 1
//!         For i = LBound(sales) To UBound(sales)
//!             range = Partition(Int(sales(i).Amount), Int(minSale), _
//!                              Int(maxSale), interval)
//!             If Not performanceBands.Exists(range) Then
//!                 performanceBands.Add range, 0
//!             End If
//!             performanceBands(range) = performanceBands(range) + 1
//!         Next i
//!         
//!         ' Generate report
//!         report = "Sales Performance Analysis" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Total Sales: " & totalSales & vbCrLf
//!         report = report & "Range: $" & Format(minSale, "#,##0") & _
//!                  " - $" & Format(maxSale, "#,##0") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         
//!         For Each range In performanceBands.Keys
//!             report = report & "Range " & range & ": " & _
//!                      performanceBands(range) & " sales (" & _
//!                      Format((performanceBands(range) / totalSales) * 100, "0.0") & _
//!                      "%)" & vbCrLf
//!         Next range
//!         
//!         AnalyzeSalesPerformance = report
//!     End Function
//!     
//!     Public Function GetPerformanceLevel(amount As Currency, _
//!                                         minAmount As Currency, _
//!                                         maxAmount As Currency) As String
//!         Dim range As String
//!         Dim interval As Currency
//!         
//!         interval = (maxAmount - minAmount) / 5
//!         range = Partition(Int(amount), Int(minAmount), Int(maxAmount), Int(interval))
//!         
//!         ' Convert to performance labels
//!         Dim parts() As String
//!         Dim lowerBound As Long
//!         parts = Split(range, ":")
//!         
//!         On Error Resume Next
//!         lowerBound = CLng(Trim(parts(0)))
//!         On Error GoTo 0
//!         
//!         If lowerBound <= minAmount + interval Then
//!             GetPerformanceLevel = "Needs Improvement"
//!         ElseIf lowerBound <= minAmount + interval * 2 Then
//!             GetPerformanceLevel = "Below Average"
//!         ElseIf lowerBound <= minAmount + interval * 3 Then
//!             GetPerformanceLevel = "Average"
//!         ElseIf lowerBound <= minAmount + interval * 4 Then
//!             GetPerformanceLevel = "Above Average"
//!         Else
//!             GetPerformanceLevel = "Excellent"
//!         End If
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Age Demographics Tool
//! ```vb
//! ' Tool for analyzing age demographics with Partition
//! Class AgeDemographicsAnalyzer
//!     Private m_ages() As Integer
//!     Private m_ageGroupSize As Integer
//!     
//!     Public Sub Initialize(ages() As Integer, groupSize As Integer)
//!         Dim i As Long
//!         ReDim m_ages(LBound(ages) To UBound(ages))
//!         For i = LBound(ages) To UBound(ages)
//!             m_ages(i) = ages(i)
//!         Next i
//!         m_ageGroupSize = groupSize
//!     End Sub
//!     
//!     Public Function GetAgeDistribution() As Object
//!         Dim dist As Object
//!         Dim i As Long
//!         Dim range As String
//!         
//!         Set dist = CreateObject("Scripting.Dictionary")
//!         
//!         For i = LBound(m_ages) To UBound(m_ages)
//!             range = Partition(m_ages(i), 0, 120, m_ageGroupSize)
//!             If Not dist.Exists(range) Then
//!                 dist.Add range, 0
//!             End If
//!             dist(range) = dist(range) + 1
//!         Next i
//!         
//!         Set GetAgeDistribution = dist
//!     End Function
//!     
//!     Public Function GetAgeGroupName(age As Integer) As String
//!         Dim range As String
//!         range = Partition(age, 0, 120, m_ageGroupSize)
//!         
//!         Select Case m_ageGroupSize
//!             Case 10
//!                 ' Decades
//!                 If InStr(range, " 0:") > 0 Then
//!                     GetAgeGroupName = "Children (0-9)"
//!                 ElseIf InStr(range, "10:") > 0 Then
//!                     GetAgeGroupName = "Teens (10-19)"
//!                 ElseIf InStr(range, "20:") > 0 Then
//!                     GetAgeGroupName = "Twenties (20-29)"
//!                 ElseIf InStr(range, "30:") > 0 Then
//!                     GetAgeGroupName = "Thirties (30-39)"
//!                 ElseIf InStr(range, "40:") > 0 Then
//!                     GetAgeGroupName = "Forties (40-49)"
//!                 ElseIf InStr(range, "50:") > 0 Then
//!                     GetAgeGroupName = "Fifties (50-59)"
//!                 ElseIf InStr(range, "60:") > 0 Then
//!                     GetAgeGroupName = "Sixties (60-69)"
//!                 Else
//!                     GetAgeGroupName = "70+"
//!                 End If
//!             Case Else
//!                 GetAgeGroupName = Trim(range)
//!         End Select
//!     End Function
//!     
//!     Public Function GenerateDemographicsReport() As String
//!         Dim report As String
//!         Dim dist As Object
//!         Dim range As Variant
//!         Dim total As Long
//!         
//!         Set dist = GetAgeDistribution()
//!         total = UBound(m_ages) - LBound(m_ages) + 1
//!         
//!         report = "Age Demographics Report" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Total Population: " & total & vbCrLf
//!         report = report & "Age Group Size: " & m_ageGroupSize & " years" & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Age Range      Count    Percentage" & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         
//!         For Each range In dist.Keys
//!             report = report & range & "  " & _
//!                      Format(dist(range), "0000") & "    " & _
//!                      Format((dist(range) / total) * 100, "00.0") & "%" & vbCrLf
//!         Next range
//!         
//!         GenerateDemographicsReport = report
//!     End Function
//!     
//!     Public Function GetMedianAgeRange() As String
//!         ' Find the range containing the median age
//!         Dim sortedAges() As Integer
//!         Dim i As Long
//!         Dim medianAge As Integer
//!         
//!         ' Copy and sort ages
//!         ReDim sortedAges(LBound(m_ages) To UBound(m_ages))
//!         For i = LBound(m_ages) To UBound(m_ages)
//!             sortedAges(i) = m_ages(i)
//!         Next i
//!         
//!         ' Simple bubble sort (for demonstration)
//!         Dim temp As Integer, j As Long
//!         For i = LBound(sortedAges) To UBound(sortedAges) - 1
//!             For j = i + 1 To UBound(sortedAges)
//!                 If sortedAges(i) > sortedAges(j) Then
//!                     temp = sortedAges(i)
//!                     sortedAges(i) = sortedAges(j)
//!                     sortedAges(j) = temp
//!                 End If
//!             Next j
//!         Next i
//!         
//!         ' Get median
//!         medianAge = sortedAges((LBound(sortedAges) + UBound(sortedAges)) \ 2)
//!         
//!         GetMedianAgeRange = Partition(medianAge, 0, 120, m_ageGroupSize)
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Test Score Grading System
//! ```vb
//! ' Automatic grading system using Partition
//! Module GradingSystem
//!     Public Function GetLetterGrade(score As Integer) As String
//!         Dim range As String
//!         
//!         ' Use Partition to determine grade range (0-100, intervals of 10)
//!         range = Partition(score, 0, 100, 10)
//!         
//!         ' Convert range to letter grade
//!         If score >= 90 Then
//!             GetLetterGrade = "A"
//!         ElseIf score >= 80 Then
//!             GetLetterGrade = "B"
//!         ElseIf score >= 70 Then
//!             GetLetterGrade = "C"
//!         ElseIf score >= 60 Then
//!             GetLetterGrade = "D"
//!         Else
//!             GetLetterGrade = "F"
//!         End If
//!     End Function
//!     
//!     Public Function AnalyzeClassScores(scores() As Integer) As String
//!         Dim gradeDistribution As Object
//!         Dim i As Long
//!         Dim range As String
//!         Dim report As String
//!         Dim total As Long
//!         
//!         Set gradeDistribution = CreateObject("Scripting.Dictionary")
//!         total = UBound(scores) - LBound(scores) + 1
//!         
//!         ' Count scores in each range
//!         For i = LBound(scores) To UBound(scores)
//!             range = Partition(scores(i), 0, 100, 10)
//!             If Not gradeDistribution.Exists(range) Then
//!                 gradeDistribution.Add range, 0
//!             End If
//!             gradeDistribution(range) = gradeDistribution(range) + 1
//!         Next i
//!         
//!         ' Build report
//!         report = "Class Score Distribution" & vbCrLf
//!         report = report & String(40, "=") & vbCrLf
//!         report = report & "Total Students: " & total & vbCrLf
//!         report = report & String(40, "-") & vbCrLf
//!         
//!         For Each range In gradeDistribution.Keys
//!             report = report & range & ": " & gradeDistribution(range) & _
//!                      " (" & Format((gradeDistribution(range) / total) * 100, "0.0") & _
//!                      "%)" & vbCrLf
//!         Next range
//!         
//!         AnalyzeClassScores = report
//!     End Function
//!     
//!     Public Function GetClassStatistics(scores() As Integer) As Object
//!         Dim stats As Object
//!         Dim dist As Object
//!         Dim i As Long
//!         Dim range As String
//!         Dim sum As Long
//!         Dim passingCount As Long
//!         
//!         Set stats = CreateObject("Scripting.Dictionary")
//!         Set dist = CreateObject("Scripting.Dictionary")
//!         
//!         sum = 0
//!         passingCount = 0
//!         
//!         For i = LBound(scores) To UBound(scores)
//!             sum = sum + scores(i)
//!             If scores(i) >= 60 Then passingCount = passingCount + 1
//!             
//!             range = Partition(scores(i), 0, 100, 10)
//!             If Not dist.Exists(range) Then dist.Add range, 0
//!             dist(range) = dist(range) + 1
//!         Next i
//!         
//!         stats.Add "Average", sum / (UBound(scores) - LBound(scores) + 1)
//!         stats.Add "PassingRate", (passingCount / (UBound(scores) - LBound(scores) + 1)) * 100
//!         stats.Add "Distribution", dist
//!         
//!         Set GetClassStatistics = stats
//!     End Function
//! End Module
//! ```
//!
//! ## Error Handling
//!
//! The `Partition` function can raise errors in the following situations:
//!
//! - **Invalid Procedure Call (Error 5)**: When:
//!   - `start` is less than 0
//!   - `stop` is less than or equal to `start`
//!   - `interval` is less than 1
//! - **Type Mismatch (Error 13)**: When arguments cannot be converted to whole numbers
//! - **Overflow (Error 6)**: When calculated values exceed data type limits
//!
//! Always validate parameters before calling `Partition`:
//!
//! ```vb
//! If start >= 0 And stop > start And interval >= 1 Then
//!     range = Partition(value, start, stop, interval)
//! Else
//!     MsgBox "Invalid partition parameters"
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - The `Partition` function is very fast for individual calls
//! - String formatting overhead is minimal but can add up with millions of calls
//! - For large datasets, consider caching partition ranges if the same parameters are used
//! - Dictionary/Collection operations for frequency counting are generally efficient
//! - Sorting large arrays for statistical analysis can be slow; consider alternative algorithms
//!
//! ## Best Practices
//!
//! 1. **Validate Parameters**: Always check that start >= 0, stop > start, and interval >= 1
//! 2. **Choose Appropriate Intervals**: Select interval sizes that create meaningful groups
//! 3. **Handle Edge Cases**: Account for values below start and above stop
//! 4. **Use Consistent Formatting**: Leverage the automatic padding for aligned displays
//! 5. **Parse Results Carefully**: Remember the format includes spaces and colons
//! 6. **Consider Alternatives**: For simple range checks, If statements may be clearer
//! 7. **Document Ranges**: Clearly document the meaning of each partition range
//! 8. **Test Boundaries**: Verify behavior at start, stop, and interval boundaries
//! 9. **Cache When Possible**: If using same parameters repeatedly, cache the range labels
//! 10. **Combine with Collections**: Use Dictionary or Collection for frequency counting
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **Partition** | Create range labels | String (range description) | Frequency distributions, histograms |
//! | **Switch** | Choose from value pairs | Variant (matched value) | Simple value mapping |
//! | **Choose** | Pick from list by index | Variant (indexed item) | Select from fixed options |
//! | **`IIf`** | Conditional expression | Variant (true/false result) | Simple binary choices |
//! | **Select Case** | Multi-way branching | N/A (statement) | Complex conditional logic |
//!
//! ## Platform and Version Notes
//!
//! - Available in VBA and VB6
//! - Behavior is consistent across Windows platforms
//! - The returned string format is fixed and cannot be customized
//! - All parameters must be whole numbers (fractional parts are truncated)
//! - Maximum values limited by Long data type (approximately ±2 billion)
//!
//! ## Limitations
//!
//! - Only works with whole numbers; fractional values are truncated
//! - Cannot customize the output format (padding, separators, etc.)
//! - Parameters must fit in Long data type range
//! - No built-in frequency counting; must implement separately
//! - String result requires parsing to extract numeric boundaries
//! - Cannot specify custom labels for ranges
//! - All ranges must have equal intervals (except first/last special cases)
//!
//! ## Related Functions
//!
//! - `Switch`: Evaluates a list of expressions and returns associated value
//! - `Choose`: Returns a value from a list based on position
//! - `IIf`: Returns one of two values based on a condition
//! - `Format`: Formats values as strings with custom patterns
//! - `InStr`: Searches for substring within a string (useful for parsing Partition results)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn partition_basic() {
        let source = r"
Dim range As String
range = Partition(15, 0, 100, 10)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("range"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("range"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("15"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
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
    fn partition_with_variables() {
        let source = r"
Dim value As Integer
Dim rangeStr As String
value = 42
rangeStr = Partition(value, 0, 100, 10)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("value"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("rangeStr"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("value"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("42"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("rangeStr"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
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
    fn partition_if_statement() {
        let source = r#"
If Partition(score, 0, 100, 10) = " 90: 99" Then
    MsgBox "Grade A"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("score"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\" 90: 99\""),
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
                        StringLiteral ("\"Grade A\""),
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
    fn partition_function_return() {
        let source = r"
Function GetAgeRange(age As Integer) As String
    GetAgeRange = Partition(age, 0, 100, 10)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetAgeRange"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("age"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetAgeRange"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Partition"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("age"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
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
    fn partition_variable_assignment() {
        let source = r"
Dim bucket As String
bucket = Partition(salesAmount, 0, 1000, 100)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("bucket"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("bucket"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("salesAmount"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("1000"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
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
    fn partition_msgbox() {
        let source = r#"
MsgBox "Value falls in range: " & Partition(num, 0, 50, 5)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Value falls in range: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Partition"),
                LeftParenthesis,
                Identifier ("num"),
                Comma,
                Whitespace,
                IntegerLiteral ("0"),
                Comma,
                Whitespace,
                IntegerLiteral ("50"),
                Comma,
                Whitespace,
                IntegerLiteral ("5"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_debug_print() {
        let source = r#"
Debug.Print "Range: " & Partition(value, min, max, interval)
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
                StringLiteral ("\"Range: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Partition"),
                LeftParenthesis,
                Identifier ("value"),
                Comma,
                Whitespace,
                Identifier ("min"),
                Comma,
                Whitespace,
                Identifier ("max"),
                Comma,
                Whitespace,
                Identifier ("interval"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_select_case() {
        let source = r#"
Select Case Partition(score, 0, 100, 20)
    Case "  0: 19"
        grade = "F"
    Case " 80: 99"
        grade = "A"
    Case Else
        grade = "Other"
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
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("score"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("20"),
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
                    StringLiteral ("\"  0: 19\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("grade"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"F\""),
                            },
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\" 80: 99\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("grade"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"A\""),
                            },
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
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("grade"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Other\""),
                            },
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
    fn partition_class_usage() {
        let source = r"
Private m_range As String

Public Sub CategorizeValue(num As Long)
    m_range = Partition(num, 0, 1000, 100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_range"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("CategorizeValue"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("num"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_range"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Partition"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("num"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1000"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
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
    fn partition_with_statement() {
        let source = r"
With analyzer
    .RangeLabel = Partition(.Value, .MinVal, .MaxVal, .Interval)
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("analyzer"),
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("RangeLabel"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Partition"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            PeriodOperator,
                                        },
                                    },
                                },
                            },
                        },
                    },
                    CallStatement {
                        Identifier ("Value"),
                        Comma,
                        Whitespace,
                        PeriodOperator,
                        Identifier ("MinVal"),
                        Comma,
                        Whitespace,
                        PeriodOperator,
                        Identifier ("MaxVal"),
                        Comma,
                        Whitespace,
                        PeriodOperator,
                        Identifier ("Interval"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_elseif() {
        let source = r#"
If x < 0 Then
    y = 1
ElseIf Partition(x, 0, 100, 10) = " 50: 59" Then
    y = 2
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("x"),
                    },
                    Whitespace,
                    LessThanOperator,
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("y"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                },
                ElseIfClause {
                    ElseIfKeyword,
                    Whitespace,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("Partition"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("x"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\" 50: 59\""),
                        },
                    },
                    Whitespace,
                    ThenKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("y"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("2"),
                            },
                            Newline,
                        },
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
    fn partition_for_loop() {
        let source = r"
For i = 0 To 100
    rangeStr = Partition(i, 0, 100, 10)
    Debug.Print i, rangeStr
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
                NumericLiteralExpression {
                    IntegerLiteral ("100"),
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("rangeStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Partition"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
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
                        Identifier ("i"),
                        Comma,
                        Whitespace,
                        Identifier ("rangeStr"),
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
    fn partition_do_while() {
        let source = r#"
Do While Partition(counter, 0, 100, 10) <> "100:   "
    counter = counter + 1
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
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    InequalityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"100:   \""),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("counter"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("counter"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
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
    fn partition_do_until() {
        let source = r#"
Do Until Partition(val, 1, 50, 5) = " 46: 50"
    val = val + 1
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                UntilKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("val"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("50"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("5"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\" 46: 50\""),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("val"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("val"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
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
    fn partition_while_wend() {
        let source = r#"
While InStr(Partition(num, 0, 1000, 100), "500") = 0
    num = num + 10
Wend
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("InStr"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Partition"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("num"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("0"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1000"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("100"),
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
                                    StringLiteral ("\"500\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("0"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("num"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("num"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                    },
                },
                WendKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_parentheses() {
        let source = r"
Dim result As String
result = (Partition(value, 0, 100, 25))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("result"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                ParenthesizedExpression {
                    LeftParenthesis,
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("25"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_iif() {
        let source = r"
Dim label As String
label = IIf(usePartition, Partition(val, 0, 100, 10), CStr(val))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("label"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("label"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("IIf"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("usePartition"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("Partition"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("val"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("CStr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("val"),
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
    fn partition_comparison() {
        let source = r#"
If Partition(val1, 0, 100, 10) = Partition(val2, 0, 100, 10) Then
    MsgBox "Same range"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("val1"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("val2"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
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
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Same range\""),
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
    fn partition_array_assignment() {
        let source = r"
Dim ranges(100) As String
ranges(i) = Partition(values(i), 0, 1000, 50)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("ranges"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("100"),
                },
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                CallExpression {
                    Identifier ("ranges"),
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
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("values"),
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
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("1000"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("50"),
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
    fn partition_property_assignment() {
        let source = r"
Set obj = New DataAnalyzer
obj.RangeBucket = Partition(obj.DataValue, 0, 500, 50)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SetStatement {
                SetKeyword,
                Whitespace,
                Identifier ("obj"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                NewKeyword,
                Whitespace,
                Identifier ("DataAnalyzer"),
                Newline,
            },
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("obj"),
                    PeriodOperator,
                    Identifier ("RangeBucket"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            MemberAccessExpression {
                                Identifier ("obj"),
                                PeriodOperator,
                                Identifier ("DataValue"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("500"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("50"),
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
    fn partition_function_argument() {
        let source = r"
Call ProcessRange(Partition(score, 0, 100, 10), studentName)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("ProcessRange"),
                LeftParenthesis,
                Identifier ("Partition"),
                LeftParenthesis,
                Identifier ("score"),
                Comma,
                Whitespace,
                IntegerLiteral ("0"),
                Comma,
                Whitespace,
                IntegerLiteral ("100"),
                Comma,
                Whitespace,
                IntegerLiteral ("10"),
                RightParenthesis,
                Comma,
                Whitespace,
                Identifier ("studentName"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_arithmetic() {
        let source = r"
Dim rangeCount As Integer
rangeCount = Len(Partition(value, 0, 100, 10))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("rangeCount"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("rangeCount"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    LenKeyword,
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Partition"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
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
    fn partition_concatenation() {
        let source = r#"
Dim msg As String
msg = "Value " & value & " is in range " & Partition(value, 0, 100, 10)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
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
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("msg"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Value \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\" is in range \""),
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("Partition"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
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
    fn partition_instr() {
        let source = r#"
Dim pos As Integer
pos = InStr(Partition(num, 0, 100, 10), ":")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("pos"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("pos"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("InStr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Partition"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("num"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
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
                                StringLiteral ("\":\""),
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
    fn partition_trim() {
        let source = r"
Dim cleaned As String
cleaned = Trim(Partition(value, 0, 1000, 100))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("cleaned"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cleaned"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Trim"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Partition"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1000"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
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
    fn partition_error_handling() {
        let source = r#"
On Error Resume Next
rangeLabel = Partition(userInput, startVal, stopVal, intervalVal)
If Err.Number <> 0 Then
    MsgBox "Invalid partition parameters"
End If
On Error GoTo 0
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
                    Identifier ("rangeLabel"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Partition"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("userInput"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("startVal"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("stopVal"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("intervalVal"),
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
                        StringLiteral ("\"Invalid partition parameters\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                GotoKeyword,
                Whitespace,
                IntegerLiteral ("0"),
                Newline,
            },
        ]);
    }

    #[test]
    fn partition_on_error_goto() {
        let source = r#"
Sub CategorizeData()
    On Error GoTo ErrorHandler
    Dim bucket As String
    bucket = Partition(dataValue, minValue, maxValue, step)
    Exit Sub
ErrorHandler:
    MsgBox "Error in partition calculation"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("CategorizeData"),
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
                        Identifier ("bucket"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bucket"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Partition"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("dataValue"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("minValue"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("maxValue"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        StepKeyword,
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
                        StringLiteral ("\"Error in partition calculation\""),
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
