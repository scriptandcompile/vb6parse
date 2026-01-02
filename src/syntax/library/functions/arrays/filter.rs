//! # Filter Function
//!
//! Returns a zero-based array containing a subset of a string array based on specified filter criteria.
//!
//! ## Syntax
//!
//! ```vb
//! Filter(sourcearray, match[, include[, compare]])
//! ```
//!
//! ## Parameters
//!
//! - **sourcearray**: Required. One-dimensional array of strings to be searched.
//! - **match**: Required. String to search for.
//! - **include**: Optional. Boolean value indicating whether to return substrings that include
//!   or exclude match. If True (default), Filter returns subset including match. If False,
//!   Filter returns subset excluding match.
//! - **compare**: Optional. Numeric value indicating the kind of string comparison to use.
//!   - 0 = vbBinaryCompare (case-sensitive, default)
//!   - 1 = vbTextCompare (case-insensitive)
//!   - 2 = vbDatabaseCompare (Microsoft Access only)
//!
//! ## Return Value
//!
//! Returns a Variant containing a zero-based array of strings. If no matches are found,
//! Filter returns an empty array. If sourcearray is Null or not a one-dimensional array,
//! an error occurs.
//!
//! ## Remarks
//!
//! The `Filter` function searches a string array for elements containing a specified substring
//! and returns a new array with matching (or non-matching) elements. This is useful for
//! filtering lists, implementing search functionality, and processing string collections.
//!
//! **Important Characteristics:**
//!
//! - Returns zero-based array regardless of input array bounds
//! - Match is substring search (not whole string match)
//! - Empty string match returns all elements (when include=True)
//! - Returns empty array if no matches found
//! - Case sensitivity controlled by compare parameter
//! - Original array is not modified
//! - Works only with one-dimensional string arrays
//! - Error 13 (Type Mismatch) if sourcearray is not an array
//! - Error 5 (Invalid procedure call) if sourcearray is multi-dimensional
//! - Error 94 (Invalid use of Null) if sourcearray is Null
//! - Returned array starts at index 0
//! - Can be used to implement NOT logic (include=False)
//!
//! ## Typical Uses
//!
//! - Filter lists based on user input
//! - Implement search functionality
//! - Remove unwanted items from arrays
//! - Find items matching a pattern
//! - Create subsets of data
//! - Filter file lists
//! - Process search results
//! - Implement autocomplete features
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim fruits() As String
//! Dim filtered() As String
//!
//! fruits = Array("Apple", "Banana", "Cherry", "Date", "Elderberry")
//!
//! ' Find fruits containing "e" (case-sensitive)
//! filtered = Filter(fruits, "e")
//! ' Returns: "Apple", "Cherry", "Date", "Elderberry"
//!
//! ' Find fruits NOT containing "e"
//! filtered = Filter(fruits, "e", False)
//! ' Returns: "Banana"
//!
//! ' Find fruits containing "a" (case-insensitive)
//! filtered = Filter(fruits, "a", True, vbTextCompare)
//! ' Returns: "Apple", "Banana", "Date"
//! ```
//!
//! ### Case-Sensitive vs Case-Insensitive
//!
//! ```vb
//! Dim names() As String
//! names = Array("John", "jane", "JAMES", "Julia", "jack")
//!
//! ' Case-sensitive search (default)
//! Dim result1() As String
//! result1 = Filter(names, "J")
//! ' Returns: "John", "JAMES", "Julia"
//!
//! ' Case-insensitive search
//! Dim result2() As String
//! result2 = Filter(names, "J", True, vbTextCompare)
//! ' Returns: "John", "jane", "JAMES", "Julia", "jack"
//! ```
//!
//! ### Exclude Matches
//!
//! ```vb
//! Dim files() As String
//! files = Array("data.txt", "backup.bak", "report.txt", "temp.bak", "notes.txt")
//!
//! ' Get only non-backup files (exclude .bak)
//! Dim textFiles() As String
//! textFiles = Filter(files, ".bak", False)
//! ' Returns: "data.txt", "report.txt", "notes.txt"
//! ```
//!
//! ## Common Patterns
//!
//! ### Filter List Based on User Input
//!
//! ```vb
//! Function SearchList(items() As String, searchTerm As String) As String()
//!     On Error GoTo ErrorHandler
//!     
//!     If Trim(searchTerm) = "" Then
//!         ' Return all items if search is empty
//!         SearchList = items
//!     Else
//!         ' Return filtered items (case-insensitive)
//!         SearchList = Filter(items, searchTerm, True, vbTextCompare)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     ' Return empty array on error
//!     Dim emptyArray() As String
//!     ReDim emptyArray(0 To -1)
//!     SearchList = emptyArray
//! End Function
//! ```
//!
//! ### Count Matching Items
//!
//! ```vb
//! Function CountMatches(items() As String, searchTerm As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     Dim matches() As String
//!     matches = Filter(items, searchTerm, True, vbTextCompare)
//!     
//!     ' Check if array is empty
//!     If UBound(matches) >= 0 Then
//!         CountMatches = UBound(matches) + 1
//!     Else
//!         CountMatches = 0
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     CountMatches = 0
//! End Function
//! ```
//!
//! ### Filter File List by Extension
//!
//! ```vb
//! Function GetFilesByExtension(files() As String, extension As String) As String()
//!     ' Ensure extension starts with dot
//!     If Left(extension, 1) <> "." Then
//!         extension = "." & extension
//!     End If
//!     
//!     ' Filter for files with this extension
//!     GetFilesByExtension = Filter(files, extension, True, vbTextCompare)
//! End Function
//!
//! ' Usage
//! Dim allFiles() As String
//! Dim txtFiles() As String
//! allFiles = Array("doc1.txt", "image.jpg", "data.txt", "photo.png")
//! txtFiles = GetFilesByExtension(allFiles, ".txt")
//! ```
//!
//! ### Multiple Filter Criteria
//!
//! ```vb
//! Function FilterMultiple(items() As String, filters() As String) As String()
//!     Dim result() As String
//!     Dim temp() As String
//!     Dim i As Long
//!     
//!     result = items
//!     
//!     ' Apply each filter sequentially
//!     For i = LBound(filters) To UBound(filters)
//!         temp = Filter(result, filters(i), True, vbTextCompare)
//!         result = temp
//!         
//!         ' Exit early if no matches
//!         If UBound(result) < 0 Then Exit For
//!     Next i
//!     
//!     FilterMultiple = result
//! End Function
//!
//! ' Usage: Find items containing both "test" and "data"
//! Dim criteria() As String
//! criteria = Array("test", "data")
//! filtered = FilterMultiple(sourceArray, criteria)
//! ```
//!
//! ### Populate `ListBox` with Filtered Results
//!
//! ```vb
//! Sub UpdateFilteredList(lstBox As ListBox, items() As String, searchText As String)
//!     Dim filtered() As String
//!     Dim i As Long
//!     
//!     lstBox.Clear
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     If Trim(searchText) = "" Then
//!         ' Show all items
//!         For i = LBound(items) To UBound(items)
//!             lstBox.AddItem items(i)
//!         Next i
//!     Else
//!         ' Show filtered items
//!         filtered = Filter(items, searchText, True, vbTextCompare)
//!         
//!         If UBound(filtered) >= 0 Then
//!             For i = 0 To UBound(filtered)
//!                 lstBox.AddItem filtered(i)
//!             Next i
//!         End If
//!     End If
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     ' Handle errors silently or show message
//! End Sub
//! ```
//!
//! ### Remove Duplicates with Filter
//!
//! ```vb
//! Function RemoveDuplicates(items() As String) As String()
//!     Dim result() As String
//!     Dim dict As Object
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     Set dict = CreateObject("Scripting.Dictionary")
//!     dict.CompareMode = vbTextCompare
//!     
//!     ' Add unique items to dictionary
//!     For i = LBound(items) To UBound(items)
//!         If Not dict.Exists(items(i)) Then
//!             dict.Add items(i), Nothing
//!         End If
//!     Next i
//!     
//!     ' Convert to array
//!     ReDim result(0 To dict.Count - 1)
//!     count = 0
//!     For i = 0 To dict.Count - 1
//!         result(count) = dict.Keys()(i)
//!         count = count + 1
//!     Next i
//!     
//!     RemoveDuplicates = result
//! End Function
//! ```
//!
//! ### Filter with Wildcard Simulation
//!
//! ```vb
//! Function FilterWildcard(items() As String, pattern As String) As Collection
//!     ' Simple wildcard: * at start, end, or both
//!     Dim results As New Collection
//!     Dim filtered() As String
//!     Dim searchTerm As String
//!     Dim i As Long
//!     Dim item As String
//!     
//!     If Left(pattern, 1) = "*" And Right(pattern, 1) = "*" Then
//!         ' Contains search
//!         searchTerm = Mid(pattern, 2, Len(pattern) - 2)
//!         filtered = Filter(items, searchTerm, True, vbTextCompare)
//!         
//!         For i = 0 To UBound(filtered)
//!             results.Add filtered(i)
//!         Next i
//!         
//!     ElseIf Left(pattern, 1) = "*" Then
//!         ' Ends with search
//!         searchTerm = Mid(pattern, 2)
//!         For i = LBound(items) To UBound(items)
//!             If Right(LCase(items(i)), Len(searchTerm)) = LCase(searchTerm) Then
//!                 results.Add items(i)
//!             End If
//!         Next i
//!         
//!     ElseIf Right(pattern, 1) = "*" Then
//!         ' Starts with search
//!         searchTerm = Left(pattern, Len(pattern) - 1)
//!         For i = LBound(items) To UBound(items)
//!             If Left(LCase(items(i)), Len(searchTerm)) = LCase(searchTerm) Then
//!                 results.Add items(i)
//!             End If
//!         Next i
//!         
//!     Else
//!         ' Exact match
//!         For i = LBound(items) To UBound(items)
//!             If LCase(items(i)) = LCase(pattern) Then
//!                 results.Add items(i)
//!             End If
//!         Next i
//!     End If
//!     
//!     Set FilterWildcard = results
//! End Function
//! ```
//!
//! ### Autocomplete Implementation
//!
//! ```vb
//! Sub TextBox_Change()
//!     Dim allItems() As String
//!     Dim matches() As String
//!     Dim i As Long
//!     
//!     ' Get all possible values (from database, array, etc.)
//!     allItems = GetAllItemNames()
//!     
//!     If Len(Me.txtSearch.Text) > 0 Then
//!         ' Filter items that start with typed text
//!         matches = Filter(allItems, Me.txtSearch.Text, True, vbTextCompare)
//!         
//!         ' Display suggestions
//!         Me.lstSuggestions.Clear
//!         
//!         If UBound(matches) >= 0 Then
//!             For i = 0 To UBound(matches)
//!                 Me.lstSuggestions.AddItem matches(i)
//!             Next i
//!             Me.lstSuggestions.Visible = True
//!         Else
//!             Me.lstSuggestions.Visible = False
//!         End If
//!     Else
//!         Me.lstSuggestions.Visible = False
//!     End If
//! End Sub
//! ```
//!
//! ### Filter Log Entries
//!
//! ```vb
//! Function FilterLogsByLevel(logEntries() As String, level As String) As String()
//!     ' Assume log format: "[LEVEL] Message"
//!     Dim levelTag As String
//!     levelTag = "[" & UCase(level) & "]"
//!     
//!     FilterLogsByLevel = Filter(logEntries, levelTag, True, vbTextCompare)
//! End Function
//!
//! ' Usage
//! Dim logs() As String
//! Dim errors() As String
//! logs = Array("[INFO] Started", "[ERROR] Failed", "[INFO] Complete", "[ERROR] Timeout")
//! errors = FilterLogsByLevel(logs, "ERROR")
//! ' Returns: "[ERROR] Failed", "[ERROR] Timeout"
//! ```
//!
//! ### Check If Array Contains Value
//!
//! ```vb
//! Function ArrayContains(items() As String, value As String, _
//!                        Optional caseSensitive As Boolean = False) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim matches() As String
//!     Dim compareMode As VbCompareMethod
//!     
//!     If caseSensitive Then
//!         compareMode = vbBinaryCompare
//!     Else
//!         compareMode = vbTextCompare
//!     End If
//!     
//!     matches = Filter(items, value, True, compareMode)
//!     
//!     ' Check if any exact matches
//!     Dim i As Long
//!     For i = 0 To UBound(matches)
//!         If StrComp(matches(i), value, compareMode) = 0 Then
//!             ArrayContains = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ArrayContains = False
//!     Exit Function
//!     
//! ErrorHandler:
//!     ArrayContains = False
//! End Function
//! ```
//!
//! ### Combine Include and Exclude Filters
//!
//! ```vb
//! Function FilterIncludeExclude(items() As String, includeText As String, _
//!                               excludeText As String) As String()
//!     Dim temp() As String
//!     
//!     ' First include items containing includeText
//!     If includeText <> "" Then
//!         temp = Filter(items, includeText, True, vbTextCompare)
//!     Else
//!         temp = items
//!     End If
//!     
//!     ' Then exclude items containing excludeText
//!     If excludeText <> "" And UBound(temp) >= 0 Then
//!         temp = Filter(temp, excludeText, False, vbTextCompare)
//!     End If
//!     
//!     FilterIncludeExclude = temp
//! End Function
//!
//! ' Usage: Get .txt files but not backup files
//! filtered = FilterIncludeExclude(files, ".txt", "backup")
//! ```
//!
//! ## Advanced Usage
//!
//! ### Dynamic Search with Multiple Columns
//!
//! ```vb
//! Type RecordData
//!     ID As String
//!     Name As String
//!     Email As String
//!     Department As String
//! End Type
//!
//! Function SearchRecords(records() As RecordData, searchTerm As String) As Long()
//!     ' Search across multiple fields and return matching indices
//!     Dim names() As String
//!     Dim emails() As String
//!     Dim departments() As String
//!     Dim matchedNames() As String
//!     Dim matchedEmails() As String
//!     Dim matchedDepts() As String
//!     Dim results() As Long
//!     Dim i As Long
//!     Dim count As Long
//!     Dim dict As Object
//!     
//!     Set dict = CreateObject("Scripting.Dictionary")
//!     
//!     ' Build arrays for each searchable field
//!     ReDim names(LBound(records) To UBound(records))
//!     ReDim emails(LBound(records) To UBound(records))
//!     ReDim departments(LBound(records) To UBound(records))
//!     
//!     For i = LBound(records) To UBound(records)
//!         names(i) = records(i).Name
//!         emails(i) = records(i).Email
//!         departments(i) = records(i).Department
//!     Next i
//!     
//!     ' Filter each field
//!     On Error Resume Next
//!     matchedNames = Filter(names, searchTerm, True, vbTextCompare)
//!     matchedEmails = Filter(emails, searchTerm, True, vbTextCompare)
//!     matchedDepts = Filter(departments, searchTerm, True, vbTextCompare)
//!     On Error GoTo 0
//!     
//!     ' Collect unique matching indices
//!     For i = LBound(records) To UBound(records)
//!         If InStr(1, records(i).Name, searchTerm, vbTextCompare) > 0 Or _
//!            InStr(1, records(i).Email, searchTerm, vbTextCompare) > 0 Or _
//!            InStr(1, records(i).Department, searchTerm, vbTextCompare) > 0 Then
//!             
//!             If Not dict.Exists(i) Then
//!                 dict.Add i, Nothing
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Convert to array
//!     If dict.Count > 0 Then
//!         ReDim results(0 To dict.Count - 1)
//!         For i = 0 To dict.Count - 1
//!             results(i) = dict.Keys()(i)
//!         Next i
//!     Else
//!         ReDim results(0 To -1)
//!     End If
//!     
//!     SearchRecords = results
//! End Function
//! ```
//!
//! ### Incremental Filter (Type-Ahead)
//!
//! ```vb
//! Private lastSearch As String
//! Private cachedResults() As String
//!
//! Sub IncrementalSearch(items() As String, currentSearch As String)
//!     Dim filtered() As String
//!     
//!     ' If new search starts with last search, filter cached results
//!     If Len(currentSearch) > Len(lastSearch) And _
//!        Left(currentSearch, Len(lastSearch)) = lastSearch And _
//!        UBound(cachedResults) >= 0 Then
//!         
//!         ' Filter from cached results (faster)
//!         filtered = Filter(cachedResults, currentSearch, True, vbTextCompare)
//!     Else
//!         ' Filter from full list
//!         filtered = Filter(items, currentSearch, True, vbTextCompare)
//!     End If
//!     
//!     ' Update cache
//!     cachedResults = filtered
//!     lastSearch = currentSearch
//!     
//!     ' Display results
//!     DisplayResults filtered
//! End Sub
//! ```
//!
//! ### Category-Based Filtering
//!
//! ```vb
//! Type Product
//!     Name As String
//!     Category As String
//!     Price As Double
//!     Description As String
//! End Type
//!
//! Function FilterProductsByCategory(products() As Product, _
//!                                   category As String) As Product()
//!     Dim categories() As String
//!     Dim filtered() As String
//!     Dim results() As Product
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     ' Build category array
//!     ReDim categories(LBound(products) To UBound(products))
//!     For i = LBound(products) To UBound(products)
//!         categories(i) = products(i).Category
//!     Next i
//!     
//!     ' Get matching categories
//!     filtered = Filter(categories, category, True, vbTextCompare)
//!     
//!     ' Build result array
//!     ReDim results(0 To UBound(filtered))
//!     count = 0
//!     
//!     For i = LBound(products) To UBound(products)
//!         If InStr(1, products(i).Category, category, vbTextCompare) > 0 Then
//!             results(count) = products(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     If count > 0 Then
//!         ReDim Preserve results(0 To count - 1)
//!     Else
//!         ReDim results(0 To -1)
//!     End If
//!     
//!     FilterProductsByCategory = results
//! End Function
//! ```
//!
//! ### Filter with Performance Tracking
//!
//! ```vb
//! Function FilterWithStats(items() As String, searchTerm As String, _
//!                          ByRef matchCount As Long, _
//!                          ByRef elapsedMs As Double) As String()
//!     Dim startTime As Double
//!     Dim results() As String
//!     
//!     startTime = Timer
//!     
//!     On Error GoTo ErrorHandler
//!     results = Filter(items, searchTerm, True, vbTextCompare)
//!     
//!     If UBound(results) >= 0 Then
//!         matchCount = UBound(results) + 1
//!     Else
//!         matchCount = 0
//!     End If
//!     
//!     elapsedMs = (Timer - startTime) * 1000
//!     
//!     FilterWithStats = results
//!     Exit Function
//!     
//! ErrorHandler:
//!     matchCount = 0
//!     elapsedMs = 0
//!     ReDim results(0 To -1)
//!     FilterWithStats = results
//! End Function
//! ```
//!
//! ### Smart Case-Sensitive Filter
//!
//! ```vb
//! Function SmartFilter(items() As String, searchTerm As String) As String()
//!     Dim compareMode As VbCompareMethod
//!     
//!     ' If search term has uppercase letters, use case-sensitive
//!     ' Otherwise use case-insensitive
//!     If searchTerm <> LCase(searchTerm) Then
//!         compareMode = vbBinaryCompare
//!     Else
//!         compareMode = vbTextCompare
//!     End If
//!     
//!     SmartFilter = Filter(items, searchTerm, True, compareMode)
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFilter(items As Variant, searchTerm As String) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     Dim emptyArray() As String
//!     
//!     ' Check if items is an array
//!     If Not IsArray(items) Then
//!         ReDim emptyArray(0 To -1)
//!         SafeFilter = emptyArray
//!         Exit Function
//!     End If
//!     
//!     ' Check if items is Null
//!     If IsNull(items) Then
//!         ReDim emptyArray(0 To -1)
//!         SafeFilter = emptyArray
//!         Exit Function
//!     End If
//!     
//!     ' Perform filter
//!     SafeFilter = Filter(items, searchTerm, True, vbTextCompare)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 13  ' Type mismatch
//!             Debug.Print "Filter error: sourcearray is not a string array"
//!         Case 5   ' Invalid procedure call
//!             Debug.Print "Filter error: sourcearray is multi-dimensional"
//!         Case 94  ' Invalid use of Null
//!             Debug.Print "Filter error: sourcearray is Null"
//!         Case Else
//!             Debug.Print "Filter error " & Err.Number & ": " & Err.Description
//!     End Select
//!     
//!     ReDim emptyArray(0 To -1)
//!     SafeFilter = emptyArray
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type Mismatch): sourcearray is not an array or not a string array
//! - **Error 5** (Invalid procedure call): sourcearray is multi-dimensional
//! - **Error 94** (Invalid use of Null): sourcearray is Null
//!
//! ## Performance Considerations
//!
//! - Filter is efficient for small to medium arrays (< 10,000 elements)
//! - For very large arrays, consider Dictionary-based approaches
//! - Case-insensitive search is slightly slower than case-sensitive
//! - Filtering already-filtered results is faster than re-filtering original array
//! - Consider caching results for repeated searches
//! - Empty string match returns entire array
//!
//! ## Best Practices
//!
//! ### Always Check Result Array
//!
//! ```vb
//! Dim results() As String
//! results = Filter(items, searchTerm)
//!
//! If UBound(results) >= 0 Then
//!     ' Process results
//!     For i = 0 To UBound(results)
//!         Debug.Print results(i)
//!     Next i
//! Else
//!     Debug.Print "No matches found"
//! End If
//! ```
//!
//! ### Use Error Handling
//!
//! ```vb
//! On Error Resume Next
//! filtered = Filter(sourceArray, searchText, True, vbTextCompare)
//! If Err.Number <> 0 Then
//!     ' Handle error
//!     ReDim filtered(0 To -1)
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Default to Case-Insensitive for User Input
//!
//! ```vb
//! ' Good - User-friendly search
//! results = Filter(items, userInput, True, vbTextCompare)
//!
//! ' Less friendly - Exact case required
//! results = Filter(items, userInput)
//! ```
//!
//! ## Comparison with Other Approaches
//!
//! ### Filter vs Manual Loop
//!
//! ```vb
//! ' Using Filter (concise)
//! matches = Filter(items, searchTerm, True, vbTextCompare)
//!
//! ' Manual loop (more control)
//! ReDim matches(0 To UBound(items))
//! count = 0
//! For i = LBound(items) To UBound(items)
//!     If InStr(1, items(i), searchTerm, vbTextCompare) > 0 Then
//!         matches(count) = items(i)
//!         count = count + 1
//!     End If
//! Next i
//! If count > 0 Then
//!     ReDim Preserve matches(0 To count - 1)
//! End If
//! ```
//!
//! ### Filter vs Collection/Dictionary
//!
//! ```vb
//! ' Filter - Returns array
//! Dim arr() As String
//! arr = Filter(items, searchTerm)
//!
//! ' Collection - More flexible but slower
//! Dim coll As New Collection
//! For i = LBound(items) To UBound(items)
//!     If InStr(1, items(i), searchTerm, vbTextCompare) > 0 Then
//!         coll.Add items(i)
//!     End If
//! Next i
//! ```
//!
//! ## Limitations
//!
//! - Works only with one-dimensional arrays
//! - Only supports string arrays
//! - Returns zero-based array (even if source is 1-based)
//! - Substring match only (no regex or wildcards)
//! - Cannot filter on multiple criteria in single call
//! - No built-in support for custom comparison functions
//! - Case-insensitive limited to vbTextCompare behavior
//!
//! ## Related Functions
//!
//! - `Array`: Creates a Variant array
//! - `Split`: Splits a string into an array
//! - `Join`: Joins array elements into a string
//! - `InStr`: Finds substring position
//! - `LBound`/`UBound`: Gets array bounds
//! - `IsArray`: Checks if variable is an array

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn filter_basic() {
        let source = r#"
filtered = Filter(fruits, "e")
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression { Identifier ("filtered") },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument { IdentifierExpression { Identifier ("fruits") } },
                        Comma,
                        Whitespace,
                        Argument { StringLiteralExpression { StringLiteral ("\"e\"") } },
                    },
                    RightParenthesis
                },
                Newline
            }
        ]);
    }

    #[test]
    fn filter_with_include() {
        let source = r#"
filtered = Filter(fruits, "e", False)
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("filtered")
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("fruits")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"e\"")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                FalseKeyword
                            }
                        }
                    },
                    RightParenthesis,
                },
                Newline,
            }
        ]);
    }

    #[test]
    fn filter_with_compare() {
        let source = r#"
filtered = Filter(fruits, "a", True, vbTextCompare)
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("filtered")
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("fruits")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"a\"")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare")
                            }
                        }
                    },
                    RightParenthesis
                },
                Newline,
            }
        ]);
    }

    #[test]
    fn filter_case_sensitive() {
        let source = r#"
result = Filter(names, "J", True, vbBinaryCompare)
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result")
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("names")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"J\"")
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword
                            }
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbBinaryCompare")
                            }
                        }
                    },
                    RightParenthesis
                },
                Newline
            },
        ]);
    }

    #[test]
    fn filter_exclude() {
        let source = r#"
textFiles = Filter(files, ".bak", False)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("textFiles"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("files"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\".bak\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                FalseKeyword,
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
    fn filter_in_function() {
        let source = r"
Function SearchList(items() As String, searchTerm As String) As String()
    SearchList = Filter(items, searchTerm, True, vbTextCompare)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("SearchList"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("items"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("searchTerm"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                LeftParenthesis,
                RightParenthesis,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("SearchList"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Filter"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("items"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("searchTerm"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbTextCompare"),
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
    fn filter_ubound_check() {
        let source = r"
matches = Filter(items, searchTerm, True, vbTextCompare)
If UBound(matches) >= 0 Then
    count = UBound(matches) + 1
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("matches"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchTerm"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
                    CallExpression {
                        Identifier ("UBound"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("matches"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOrEqualOperator,
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
                            Identifier ("count"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("UBound"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("matches"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_sequential() {
        let source = r"
temp = Filter(result, filters(i), True, vbTextCompare)
result = temp
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("result"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("filters"),
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
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_in_loop() {
        let source = r#"
filtered = Filter(sourceArray, "test")
For i = 0 To UBound(filtered)
    lstBox.AddItem filtered(i)
Next i
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("filtered"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("sourceArray"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"test\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
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
                CallExpression {
                    Identifier ("UBound"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("filtered"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("lstBox"),
                        PeriodOperator,
                        Identifier ("AddItem"),
                        Whitespace,
                        Identifier ("filtered"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
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
    fn filter_error_handling() {
        let source = r"
On Error GoTo ErrorHandler
filtered = Filter(sourceArray, searchText, True, vbTextCompare)
Exit Function
ErrorHandler:
    ReDim filtered(0 To -1)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                GotoKeyword,
                Whitespace,
                Identifier ("ErrorHandler"),
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("filtered"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("sourceArray"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchText"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            ExitStatement {
                ExitKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
            LabelStatement {
                Identifier ("ErrorHandler"),
                ColonOperator,
                Newline,
            },
            Whitespace,
            ReDimStatement {
                ReDimKeyword,
                Whitespace,
                Identifier ("filtered"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("0"),
                },
                Whitespace,
                ToKeyword,
                Whitespace,
                UnaryExpression {
                    SubtractionOperator,
                    NumericLiteralExpression {
                        IntegerLiteral ("1"),
                    },
                },
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_if_statement() {
        let source = r#"
If Trim(searchTerm) = "" Then
    results = items
Else
    results = Filter(items, searchTerm, True, vbTextCompare)
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
                        Identifier ("Trim"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("searchTerm"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("results"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("items"),
                        },
                        Newline,
                    },
                },
                ElseClause {
                    ElseKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("results"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Filter"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("items"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("searchTerm"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            TrueKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbTextCompare"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
    fn filter_extension() {
        let source = r#"
txtFiles = Filter(allFiles, ".txt", True, vbTextCompare)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("txtFiles"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("allFiles"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\".txt\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
    fn filter_with_variables() {
        let source = r"
searchMode = vbTextCompare
includeMatches = True
result = Filter(sourceData, pattern, includeMatches, searchMode)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("searchMode"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                IdentifierExpression {
                    Identifier ("vbTextCompare"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("includeMatches"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BooleanLiteralExpression {
                    TrueKeyword,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("sourceData"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("pattern"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("includeMatches"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchMode"),
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
    fn filter_concatenation() {
        let source = r#"
levelTag = "[" & UCase(level) & "]"
result = Filter(logEntries, levelTag, True, vbTextCompare)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("levelTag"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        StringLiteralExpression {
                            StringLiteral ("\"[\""),
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        CallExpression {
                            Identifier ("UCase"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("level"),
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
                        StringLiteral ("\"]\""),
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("logEntries"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("levelTag"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
    fn filter_combine_operations() {
        let source = r"
temp = Filter(items, includeText, True, vbTextCompare)
temp = Filter(temp, excludeText, False, vbTextCompare)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("includeText"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("temp"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("excludeText"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                FalseKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
    fn filter_debug_print() {
        let source = r#"
Debug.Print "Found: " & UBound(Filter(items, searchTerm)) + 1
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
                StringLiteral ("\"Found: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("UBound"),
                LeftParenthesis,
                Identifier ("Filter"),
                LeftParenthesis,
                Identifier ("items"),
                Comma,
                Whitespace,
                Identifier ("searchTerm"),
                RightParenthesis,
                RightParenthesis,
                Whitespace,
                AdditionOperator,
                Whitespace,
                IntegerLiteral ("1"),
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_array_access() {
        let source = r"
matchedNames = Filter(names, searchTerm, True, vbTextCompare)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("matchedNames"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("names"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchTerm"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
    fn filter_instr_comparison() {
        let source = r"
If InStr(1, records(i).Name, searchTerm, vbTextCompare) > 0 Then
    matches = Filter(names, searchTerm, True, vbTextCompare)
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
                        Identifier ("InStr"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                MemberAccessExpression {
                                    CallExpression {
                                        Identifier ("records"),
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
                                    NameKeyword,
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    Identifier ("searchTerm"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    Identifier ("vbTextCompare"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
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
                            Identifier ("matches"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Filter"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("names"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("searchTerm"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbTextCompare"),
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
    fn filter_cache_update() {
        let source = r"
filtered = Filter(cachedResults, currentSearch, True, vbTextCompare)
cachedResults = filtered
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("filtered"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("cachedResults"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("currentSearch"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cachedResults"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                IdentifierExpression {
                    Identifier ("filtered"),
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_empty_check() {
        let source = r#"
results = Filter(items, searchTerm, True, vbTextCompare)
If UBound(results) < 0 Then
    Debug.Print "No matches"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("results"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchTerm"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
                    CallExpression {
                        Identifier ("UBound"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("results"),
                                },
                            },
                        },
                        RightParenthesis,
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
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"No matches\""),
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
    fn filter_smart_case() {
        let source = r"
If searchTerm <> LCase(searchTerm) Then
    compareMode = vbBinaryCompare
Else
    compareMode = vbTextCompare
End If
result = Filter(items, searchTerm, True, compareMode)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("searchTerm"),
                    },
                    Whitespace,
                    InequalityOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("LCase"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("searchTerm"),
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
                            Identifier ("compareMode"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("vbBinaryCompare"),
                        },
                        Newline,
                    },
                },
                ElseClause {
                    ElseKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("compareMode"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
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
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchTerm"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("compareMode"),
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
    fn filter_isnull_check() {
        let source = r"
If Not IsNull(items) Then
    result = Filter(items, searchTerm, True, vbTextCompare)
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                UnaryExpression {
                    NotKeyword,
                    Whitespace,
                    CallExpression {
                        Identifier ("IsNull"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("items"),
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
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Filter"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("items"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("searchTerm"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbTextCompare"),
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
    fn filter_isarray_check() {
        let source = r"
If IsArray(items) Then
    filtered = Filter(items, searchTerm, True, vbTextCompare)
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("IsArray"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("filtered"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Filter"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("items"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("searchTerm"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbTextCompare"),
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
    fn filter_timer_tracking() {
        let source = r"
startTime = Timer
results = Filter(items, searchTerm, True, vbTextCompare)
elapsedMs = (Timer - startTime) * 1000
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("startTime"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                IdentifierExpression {
                    Identifier ("Timer"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("results"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("items"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("searchTerm"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("elapsedMs"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("Timer"),
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("startTime"),
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("1000"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_for_each_prep() {
        let source = r"
Dim item As Variant
matches = Filter(sourceArray, pattern, True, vbTextCompare)
For Each item In matches
    Debug.Print item
Next item
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("item"),
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("matches"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Filter"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("sourceArray"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("pattern"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbTextCompare"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            ForEachStatement {
                ForKeyword,
                Whitespace,
                EachKeyword,
                Whitespace,
                Identifier ("item"),
                Whitespace,
                InKeyword,
                Whitespace,
                Identifier ("matches"),
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("item"),
                        Newline,
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("item"),
                Newline,
            },
        ]);
    }

    #[test]
    fn filter_select_case() {
        let source = r#"
Select Case filterType
    Case "include"
        result = Filter(items, term, True, vbTextCompare)
    Case "exclude"
        result = Filter(items, term, False, vbTextCompare)
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
                IdentifierExpression {
                    Identifier ("filterType"),
                },
                Newline,
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\"include\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("result"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Filter"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("items"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("term"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            TrueKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbTextCompare"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\"exclude\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("result"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Filter"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("items"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("term"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            FalseKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbTextCompare"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
}
