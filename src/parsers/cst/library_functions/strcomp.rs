//! VB6 `StrComp` Function
//!
//! The `StrComp` function compares two strings and returns a value indicating their relationship.
//!
//! ## Syntax
//! ```vb6
//! StrComp(string1, string2[, compare])
//! ```
//!
//! ## Parameters
//! - `string1`: Required. Any valid string expression.
//! - `string2`: Required. Any valid string expression.
//! - `compare`: Optional. Specifies the type of string comparison. Can be one of the following constants:
//!   - `vbBinaryCompare` (0): Performs a binary comparison (case-sensitive, based on character codes)
//!   - `vbTextCompare` (1): Performs a textual comparison (case-insensitive)
//!   - `vbDatabaseCompare` (2): Performs a comparison based on database information (Microsoft Access only)
//!
//! ## Returns
//! Returns an `Integer` indicating the result of the comparison:
//! - `-1` if `string1` is less than `string2`
//! - `0` if `string1` equals `string2`
//! - `1` if `string1` is greater than `string2`
//! - `Null` if either `string1` or `string2` is `Null`
//!
//! ## Remarks
//! The `StrComp` function provides flexible string comparison with control over case sensitivity:
//!
//! - **Binary comparison (vbBinaryCompare = 0)**: Compares strings based on internal binary representation (character codes). This is case-sensitive and faster. "A" < "a" because uppercase letters have lower ASCII values than lowercase letters.
//! - **Text comparison (vbTextCompare = 1)**: Compares strings in a case-insensitive manner. "A" = "a" in text comparison.
//! - **Default behavior**: If `compare` argument is omitted, the comparison mode is determined by the `Option Compare` statement at the module level. If no `Option Compare` is specified, binary comparison is used.
//! - **Null handling**: If either string is `Null`, the function returns `Null` (not an error).
//! - **Empty strings**: Empty strings ("") are less than any non-empty string.
//! - **Comparison logic**: Uses lexicographic ordering based on character codes (binary) or case-folded characters (text).
//! - **Performance**: Binary comparison is faster than text comparison.
//! - **Unicode support**: VB6 uses Unicode internally, so comparisons work correctly with international characters.
//!
//! ### Option Compare Statement
//! The module-level `Option Compare` statement affects default comparison behavior:
//! ```vb6
//! Option Compare Binary   ' Default: case-sensitive comparisons
//! Option Compare Text     ' Case-insensitive comparisons
//! Option Compare Database ' Database-based comparisons (Access only)
//! ```
//!
//! ### Comparison with Operators
//! - **`StrComp` vs = operator**: The `=` operator returns Boolean (True/False), while `StrComp` returns Integer (-1, 0, 1) providing ordering information.
//! - **`StrComp` vs `InStr`**: `InStr` finds substring position, while `StrComp` compares entire strings.
//! - **`StrComp` vs Like**: `Like` supports pattern matching with wildcards, while `StrComp` is exact comparison.
//!
//! ## Typical Uses
//! 1. **Sorting Algorithms**: Implement custom sorting routines for string arrays
//! 2. **Case-Insensitive Comparisons**: Compare strings without regard to case
//! 3. **Data Validation**: Verify string equality with specific comparison rules
//! 4. **Search Operations**: Find matching strings in collections with case control
//! 5. **Alphabetical Ordering**: Determine alphabetical order for display or reports
//! 6. **User Input Validation**: Compare user input against expected values case-insensitively
//! 7. **Database Queries**: Build comparison logic for filtering and searching
//! 8. **File Comparisons**: Compare filenames or paths with case sensitivity control
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic String Comparison
//! ```vb6
//! Dim result As Integer
//!
//! ' Binary comparison (case-sensitive)
//! result = StrComp("ABC", "abc", vbBinaryCompare)   ' Returns -1 (ABC < abc)
//! result = StrComp("ABC", "ABC", vbBinaryCompare)   ' Returns 0 (equal)
//! result = StrComp("abc", "ABC", vbBinaryCompare)   ' Returns 1 (abc > ABC)
//!
//! ' Text comparison (case-insensitive)
//! result = StrComp("ABC", "abc", vbTextCompare)     ' Returns 0 (equal)
//! result = StrComp("ABC", "ABD", vbTextCompare)     ' Returns -1 (ABC < ABD)
//! result = StrComp("XYZ", "ABC", vbTextCompare)     ' Returns 1 (XYZ > ABC)
//! ```
//!
//! ### Example 2: Using Return Value
//! ```vb6
//! Dim str1 As String
//! Dim str2 As String
//! Dim compareResult As Integer
//!
//! str1 = "Apple"
//! str2 = "apple"
//!
//! compareResult = StrComp(str1, str2, vbTextCompare)
//!
//! Select Case compareResult
//!     Case -1
//!         MsgBox str1 & " comes before " & str2
//!     Case 0
//!         MsgBox str1 & " equals " & str2
//!     Case 1
//!         MsgBox str1 & " comes after " & str2
//! End Select
//! ```
//!
//! ### Example 3: Case-Insensitive Search
//! ```vb6
//! Function FindInArray(arr() As String, searchValue As String) As Integer
//!     Dim i As Integer
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If StrComp(arr(i), searchValue, vbTextCompare) = 0 Then
//!             FindInArray = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     FindInArray = -1  ' Not found
//! End Function
//! ```
//!
//! ### Example 4: Null Handling
//! ```vb6
//! Dim result As Variant
//! Dim str1 As Variant
//! Dim str2 As Variant
//!
//! str1 = "Hello"
//! str2 = Null
//!
//! result = StrComp(str1, str2)  ' Returns Null
//!
//! If IsNull(result) Then
//!     MsgBox "One of the strings is Null"
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Bubble Sort with `StrComp`
//! ```vb6
//! Sub SortStrings(arr() As String, caseInsensitive As Boolean)
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim temp As String
//!     Dim compareMode As Integer
//!     
//!     compareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
//!     
//!     For i = LBound(arr) To UBound(arr) - 1
//!         For j = i + 1 To UBound(arr)
//!             If StrComp(arr(i), arr(j), compareMode) > 0 Then
//!                 temp = arr(i)
//!                 arr(i) = arr(j)
//!                 arr(j) = temp
//!             End If
//!         Next j
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 2: Case-Insensitive Equality Check
//! ```vb6
//! Function EqualsIgnoreCase(str1 As String, str2 As String) As Boolean
//!     EqualsIgnoreCase = (StrComp(str1, str2, vbTextCompare) = 0)
//! End Function
//! ```
//!
//! ### Pattern 3: Find Minimum String
//! ```vb6
//! Function FindMinString(arr() As String) As String
//!     Dim i As Integer
//!     Dim minStr As String
//!     
//!     If UBound(arr) < LBound(arr) Then Exit Function
//!     
//!     minStr = arr(LBound(arr))
//!     For i = LBound(arr) + 1 To UBound(arr)
//!         If StrComp(arr(i), minStr, vbTextCompare) < 0 Then
//!             minStr = arr(i)
//!         End If
//!     Next i
//!     
//!     FindMinString = minStr
//! End Function
//! ```
//!
//! ### Pattern 4: Binary Search (Sorted Array)
//! ```vb6
//! Function BinarySearch(arr() As String, searchValue As String) As Integer
//!     Dim low As Integer
//!     Dim high As Integer
//!     Dim mid As Integer
//!     Dim compareResult As Integer
//!     
//!     low = LBound(arr)
//!     high = UBound(arr)
//!     
//!     Do While low <= high
//!         mid = (low + high) \ 2
//!         compareResult = StrComp(arr(mid), searchValue, vbTextCompare)
//!         
//!         If compareResult = 0 Then
//!             BinarySearch = mid
//!             Exit Function
//!         ElseIf compareResult < 0 Then
//!             low = mid + 1
//!         Else
//!             high = mid - 1
//!         End If
//!     Loop
//!     
//!     BinarySearch = -1  ' Not found
//! End Function
//! ```
//!
//! ### Pattern 5: Validate Against List
//! ```vb6
//! Function IsValidValue(value As String, validValues() As String) As Boolean
//!     Dim i As Integer
//!     
//!     For i = LBound(validValues) To UBound(validValues)
//!         If StrComp(value, validValues(i), vbTextCompare) = 0 Then
//!             IsValidValue = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsValidValue = False
//! End Function
//! ```
//!
//! ### Pattern 6: Remove Duplicates (Case-Insensitive)
//! ```vb6
//! Function RemoveDuplicates(arr() As String) As String()
//!     Dim result() As String
//!     Dim count As Integer
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim isDuplicate As Boolean
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         isDuplicate = False
//!         For j = 0 To count - 1
//!             If StrComp(arr(i), result(j), vbTextCompare) = 0 Then
//!                 isDuplicate = True
//!                 Exit For
//!             End If
//!         Next j
//!         
//!         If Not isDuplicate Then
//!             ReDim Preserve result(0 To count)
//!             result(count) = arr(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     RemoveDuplicates = result
//! End Function
//! ```
//!
//! ### Pattern 7: Group By First Letter
//! ```vb6
//! Function GroupByFirstLetter(arr() As String) As Collection
//!     Dim groups As New Collection
//!     Dim i As Integer
//!     Dim firstLetter As String
//!     Dim group As Collection
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         firstLetter = UCase$(Left$(arr(i), 1))
//!         
//!         On Error Resume Next
//!         Set group = groups(firstLetter)
//!         On Error GoTo 0
//!         
//!         If group Is Nothing Then
//!             Set group = New Collection
//!             groups.Add group, firstLetter
//!         End If
//!         
//!         group.Add arr(i)
//!         Set group = Nothing
//!     Next i
//!     
//!     Set GroupByFirstLetter = groups
//! End Function
//! ```
//!
//! ### Pattern 8: Natural Sort Helper
//! ```vb6
//! Function CompareNatural(str1 As String, str2 As String) As Integer
//!     ' Simple natural sort: compare non-numeric parts with StrComp
//!     ' This is a simplified version
//!     
//!     If IsNumeric(str1) And IsNumeric(str2) Then
//!         If CDbl(str1) < CDbl(str2) Then
//!             CompareNatural = -1
//!         ElseIf CDbl(str1) > CDbl(str2) Then
//!             CompareNatural = 1
//!         Else
//!             CompareNatural = 0
//!         End If
//!     Else
//!         CompareNatural = StrComp(str1, str2, vbTextCompare)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: Case-Sensitive Contains
//! ```vb6
//! Function ContainsCaseSensitive(arr() As String, value As String) As Boolean
//!     Dim i As Integer
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If StrComp(arr(i), value, vbBinaryCompare) = 0 Then
//!             ContainsCaseSensitive = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ContainsCaseSensitive = False
//! End Function
//! ```
//!
//! ### Pattern 10: Compare File Extensions
//! ```vb6
//! Function HasExtension(filename As String, extension As String) As Boolean
//!     Dim fileExt As String
//!     Dim dotPos As Integer
//!     
//!     dotPos = InStrRev(filename, ".")
//!     If dotPos = 0 Then
//!         HasExtension = False
//!         Exit Function
//!     End If
//!     
//!     fileExt = Mid$(filename, dotPos + 1)
//!     HasExtension = (StrComp(fileExt, extension, vbTextCompare) = 0)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: String Sorter Class
//! ```vb6
//! ' Class: StringSorter
//! ' Provides sorting capabilities with various comparison modes
//! Option Explicit
//!
//! Public Enum SortOrder
//!     Ascending = 1
//!     Descending = -1
//! End Enum
//!
//! Private m_CompareMode As VbCompareMethod
//! Private m_SortOrder As SortOrder
//!
//! Public Sub Initialize(Optional compareMode As VbCompareMethod = vbBinaryCompare, _
//!                       Optional sortOrder As SortOrder = Ascending)
//!     m_CompareMode = compareMode
//!     m_SortOrder = sortOrder
//! End Sub
//!
//! Public Sub QuickSort(arr() As String, Optional leftIndex As Long = -1, _
//!                                       Optional rightIndex As Long = -1)
//!     Dim i As Long
//!     Dim j As Long
//!     Dim pivot As String
//!     Dim temp As String
//!     
//!     ' Initialize indices on first call
//!     If leftIndex = -1 Then leftIndex = LBound(arr)
//!     If rightIndex = -1 Then rightIndex = UBound(arr)
//!     
//!     If leftIndex >= rightIndex Then Exit Sub
//!     
//!     ' Choose pivot
//!     pivot = arr((leftIndex + rightIndex) \ 2)
//!     i = leftIndex
//!     j = rightIndex
//!     
//!     ' Partition
//!     Do While i <= j
//!         Do While CompareValues(arr(i), pivot) < 0
//!             i = i + 1
//!         Loop
//!         
//!         Do While CompareValues(arr(j), pivot) > 0
//!             j = j - 1
//!         Loop
//!         
//!         If i <= j Then
//!             temp = arr(i)
//!             arr(i) = arr(j)
//!             arr(j) = temp
//!             i = i + 1
//!             j = j - 1
//!         End If
//!     Loop
//!     
//!     ' Recursive calls
//!     If leftIndex < j Then QuickSort arr, leftIndex, j
//!     If i < rightIndex Then QuickSort arr, i, rightIndex
//! End Sub
//!
//! Private Function CompareValues(str1 As String, str2 As String) As Integer
//!     CompareValues = StrComp(str1, str2, m_CompareMode) * m_SortOrder
//! End Function
//!
//! Public Function IsSorted(arr() As String) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(arr) To UBound(arr) - 1
//!         If CompareValues(arr(i), arr(i + 1)) > 0 Then
//!             IsSorted = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsSorted = True
//! End Function
//! ```
//!
//! ### Example 2: Dictionary with Case-Insensitive Keys
//! ```vb6
//! ' Class: CaseInsensitiveDictionary
//! ' Dictionary that treats keys as case-insensitive
//! Option Explicit
//!
//! Private m_Keys() As String
//! Private m_Values() As Variant
//! Private m_Count As Long
//!
//! Public Sub Initialize()
//!     m_Count = 0
//!     ReDim m_Keys(0 To 9)
//!     ReDim m_Values(0 To 9)
//! End Sub
//!
//! Public Sub Add(key As String, value As Variant)
//!     Dim index As Long
//!     
//!     ' Check if key already exists
//!     index = FindKey(key)
//!     If index >= 0 Then
//!         Err.Raise 457, , "Key already exists"
//!     End If
//!     
//!     ' Resize if necessary
//!     If m_Count > UBound(m_Keys) Then
//!         ReDim Preserve m_Keys(0 To UBound(m_Keys) * 2 + 1)
//!         ReDim Preserve m_Values(0 To UBound(m_Values) * 2 + 1)
//!     End If
//!     
//!     ' Add new item
//!     m_Keys(m_Count) = key
//!     If IsObject(value) Then
//!         Set m_Values(m_Count) = value
//!     Else
//!         m_Values(m_Count) = value
//!     End If
//!     m_Count = m_Count + 1
//! End Sub
//!
//! Public Function Item(key As String) As Variant
//!     Dim index As Long
//!     index = FindKey(key)
//!     
//!     If index < 0 Then
//!         Err.Raise 5, , "Key not found"
//!     End If
//!     
//!     If IsObject(m_Values(index)) Then
//!         Set Item = m_Values(index)
//!     Else
//!         Item = m_Values(index)
//!     End If
//! End Function
//!
//! Public Function Exists(key As String) As Boolean
//!     Exists = (FindKey(key) >= 0)
//! End Function
//!
//! Private Function FindKey(key As String) As Long
//!     Dim i As Long
//!     
//!     For i = 0 To m_Count - 1
//!         If StrComp(m_Keys(i), key, vbTextCompare) = 0 Then
//!             FindKey = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     FindKey = -1
//! End Function
//!
//! Public Property Get Count() As Long
//!     Count = m_Count
//! End Property
//! ```
//!
//! ### Example 3: String Matcher Module
//! ```vb6
//! ' Module: StringMatcher
//! ' Advanced string matching and searching utilities
//! Option Explicit
//!
//! Public Function FindAllMatches(arr() As String, searchValue As String, _
//!                                caseInsensitive As Boolean) As Long()
//!     Dim matches() As Long
//!     Dim matchCount As Long
//!     Dim i As Long
//!     Dim compareMode As VbCompareMethod
//!     
//!     compareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
//!     matchCount = 0
//!     ReDim matches(0 To UBound(arr) - LBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If StrComp(arr(i), searchValue, compareMode) = 0 Then
//!             matches(matchCount) = i
//!             matchCount = matchCount + 1
//!         End If
//!     Next i
//!     
//!     ' Resize to actual count
//!     If matchCount > 0 Then
//!         ReDim Preserve matches(0 To matchCount - 1)
//!     Else
//!         ReDim matches(0 To -1)  ' Empty array
//!     End If
//!     
//!     FindAllMatches = matches
//! End Function
//!
//! Public Function GetUniqueValues(arr() As String, caseInsensitive As Boolean) As String()
//!     Dim unique() As String
//!     Dim uniqueCount As Long
//!     Dim i As Long
//!     Dim j As Long
//!     Dim found As Boolean
//!     Dim compareMode As VbCompareMethod
//!     
//!     compareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
//!     uniqueCount = 0
//!     ReDim unique(0 To UBound(arr) - LBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         found = False
//!         For j = 0 To uniqueCount - 1
//!             If StrComp(arr(i), unique(j), compareMode) = 0 Then
//!                 found = True
//!                 Exit For
//!             End If
//!         Next j
//!         
//!         If Not found Then
//!             unique(uniqueCount) = arr(i)
//!             uniqueCount = uniqueCount + 1
//!         End If
//!     Next i
//!     
//!     ' Resize to actual count
//!     If uniqueCount > 0 Then
//!         ReDim Preserve unique(0 To uniqueCount - 1)
//!     Else
//!         ReDim unique(0 To -1)
//!     End If
//!     
//!     GetUniqueValues = unique
//! End Function
//!
//! Public Function CompareArrays(arr1() As String, arr2() As String, _
//!                               caseInsensitive As Boolean) As Boolean
//!     Dim i As Long
//!     Dim compareMode As VbCompareMethod
//!     
//!     ' Check bounds
//!     If UBound(arr1) - LBound(arr1) <> UBound(arr2) - LBound(arr2) Then
//!         CompareArrays = False
//!         Exit Function
//!     End If
//!     
//!     compareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
//!     
//!     ' Compare elements
//!     For i = 0 To UBound(arr1) - LBound(arr1)
//!         If StrComp(arr1(LBound(arr1) + i), arr2(LBound(arr2) + i), compareMode) <> 0 Then
//!             CompareArrays = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     CompareArrays = True
//! End Function
//!
//! Public Function CountOccurrences(arr() As String, searchValue As String, _
//!                                  caseInsensitive As Boolean) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     Dim compareMode As VbCompareMethod
//!     
//!     compareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
//!     count = 0
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If StrComp(arr(i), searchValue, compareMode) = 0 Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountOccurrences = count
//! End Function
//! ```
//!
//! ### Example 4: Validation Helper Class
//! ```vb6
//! ' Class: StringValidator
//! ' Provides string validation with comparison options
//! Option Explicit
//!
//! Private m_ValidValues As Collection
//! Private m_CaseInsensitive As Boolean
//!
//! Public Sub Initialize(caseInsensitive As Boolean)
//!     Set m_ValidValues = New Collection
//!     m_CaseInsensitive = caseInsensitive
//! End Sub
//!
//! Public Sub AddValidValue(value As String)
//!     ' Check for duplicates
//!     If Not IsValid(value) Then
//!         m_ValidValues.Add value
//!     End If
//! End Sub
//!
//! Public Function IsValid(value As String) As Boolean
//!     Dim item As Variant
//!     Dim compareMode As VbCompareMethod
//!     
//!     compareMode = IIf(m_CaseInsensitive, vbTextCompare, vbBinaryCompare)
//!     
//!     For Each item In m_ValidValues
//!         If StrComp(CStr(item), value, compareMode) = 0 Then
//!             IsValid = True
//!             Exit Function
//!         End If
//!     Next item
//!     
//!     IsValid = False
//! End Function
//!
//! Public Function GetClosestMatch(value As String) As String
//!     Dim item As Variant
//!     Dim minDistance As Long
//!     Dim distance As Long
//!     Dim closest As String
//!     
//!     minDistance = 999999
//!     
//!     For Each item In m_ValidValues
//!         distance = GetEditDistance(value, CStr(item))
//!         If distance < minDistance Then
//!             minDistance = distance
//!             closest = CStr(item)
//!         End If
//!     Next item
//!     
//!     GetClosestMatch = closest
//! End Function
//!
//! Private Function GetEditDistance(str1 As String, str2 As String) As Long
//!     ' Simple implementation - just return length difference
//!     GetEditDistance = Abs(Len(str1) - Len(str2))
//! End Function
//!
//! Public Sub Clear()
//!     Set m_ValidValues = New Collection
//! End Sub
//!
//! Public Property Get Count() As Long
//!     Count = m_ValidValues.Count
//! End Property
//! ```
//!
//! ## Error Handling
//! The `StrComp` function does not raise errors under normal circumstances. However:
//!
//! - Returns `Null` if either string argument is `Null` (not an error)
//! - **Error 13 (Type mismatch)**: If arguments cannot be converted to strings
//! - **Error 5 (Invalid procedure call)**: If `compare` argument is not 0, 1, or 2
//!
//! ## Performance Notes
//! - **Binary comparison** is faster than text comparison
//! - Very fast for short strings (< 100 characters)
//! - Performance scales linearly with string length
//! - Text comparison requires case folding, adding overhead
//! - Consider using `=` operator if simple equality check suffices
//! - Cache comparison mode constant if used repeatedly in loops
//!
//! ## Best Practices
//! 1. **Use vbTextCompare constants** instead of numeric values (0, 1, 2) for clarity
//! 2. **Set Option Compare** at module level for consistent default behavior
//! 3. **Handle Null values** explicitly with `IsNull` check when dealing with Variants
//! 4. **Choose appropriate mode**: Use binary for exact matching, text for user-facing comparisons
//! 5. **Cache compare mode** in variables when using the same mode repeatedly
//! 6. **Use return value properly**: Remember -1, 0, 1 (not True/False)
//! 7. **Consider = operator** for simple equality checks (may be optimized better)
//! 8. **Document comparison mode** in function signatures and comments
//! 9. **Test edge cases**: Empty strings, Null values, Unicode characters
//! 10. **Use for sorting**: `StrComp` is ideal for custom sort implementations
//!
//! ## Comparison Table
//!
//! | Function | Returns | Case-Sensitive | Ordering | Null Handling |
//! |----------|---------|----------------|----------|---------------|
//! | `StrComp` | -1/0/1 | Configurable | Yes | Returns Null |
//! | `=` operator | Boolean | Per Option Compare | No | Error if Null |
//! | `InStr` | Position | Configurable | No | Returns Null |
//! | `Like` | Boolean | Per Option Compare | No | False if Null |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Behavior consistent across platforms
//! - `vbDatabaseCompare` only meaningful in Microsoft Access
//! - Unicode comparison works correctly with international characters
//! - Option Compare affects default behavior differently in VBA vs VB6
//!
//! ## Limitations
//! - No support for locale-specific collation beyond text/binary
//! - Cannot specify custom comparison rules
//! - `vbDatabaseCompare` mode rarely useful outside Access
//! - Null handling returns Null rather than error (can be unexpected)
//! - No indication of *where* strings differ, only that they do
//! - Cannot compare string arrays directly (must loop)
//! - No natural sort support (e.g., "file2.txt" vs "file10.txt")

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn strcomp_basic() {
        let source = r#"
Sub Test()
    result = StrComp("ABC", "abc", vbBinaryCompare)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
        assert!(debug.contains("ABC"));
    }

    #[test]
    fn strcomp_variable_assignment() {
        let source = r#"
Sub Test()
    Dim result As Integer
    result = StrComp(str1, str2)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
        assert!(debug.contains("str1"));
    }

    #[test]
    fn strcomp_text_compare() {
        let source = r#"
Sub Test()
    result = StrComp(name1, name2, vbTextCompare)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
        assert!(debug.contains("vbTextCompare"));
    }

    #[test]
    fn strcomp_if_statement() {
        let source = r#"
Sub Test()
    If StrComp(str1, str2, vbTextCompare) = 0 Then
        MsgBox "Equal"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_select_case() {
        let source = r#"
Sub Test()
    Select Case StrComp(str1, str2)
        Case -1
            MsgBox "Less than"
        Case 0
            MsgBox "Equal"
        Case 1
            MsgBox "Greater than"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_for_loop() {
        let source = r#"
Sub Test()
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), searchValue, vbTextCompare) = 0 Then
            Exit For
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_function_return() {
        let source = r#"
Function Compare(s1 As String, s2 As String) As Integer
    Compare = StrComp(s1, s2, vbTextCompare)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_do_while() {
        let source = r#"
Sub Test()
    Do While StrComp(current, target, vbTextCompare) <> 0
        current = GetNext()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_sorting() {
        let source = r#"
Sub Test()
    If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_comparison() {
        let source = r#"
Sub Test()
    Dim isLess As Boolean
    isLess = (StrComp(str1, str2) < 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessResult(StrComp(a, b, vbTextCompare))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_iif() {
        let source = r#"
Sub Test()
    result = IIf(StrComp(str1, str2, vbTextCompare) = 0, "Same", "Different")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_array_search() {
        let source = r#"
Sub Test()
    For Each item In collection
        If StrComp(item, searchTerm, vbTextCompare) = 0 Then
            found = True
        End If
    Next
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_binary_compare() {
        let source = r#"
Sub Test()
    result = StrComp("Test", "test", vbBinaryCompare)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
        assert!(debug.contains("vbBinaryCompare"));
    }

    #[test]
    fn strcomp_while_wend() {
        let source = r#"
Sub Test()
    While StrComp(str1, str2, vbTextCompare) <> 0
        str1 = Modify(str1)
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_do_until() {
        let source = r#"
Sub Test()
    Do Until StrComp(input, expected, vbTextCompare) = 0
        input = GetInput()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_with_statement() {
        let source = r#"
Sub Test()
    With obj
        If StrComp(.Name, targetName, vbTextCompare) = 0 Then
            found = True
        End If
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_parentheses() {
        let source = r#"
Sub Test()
    result = (StrComp(str1, str2, vbTextCompare) = 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = StrComp(var1, var2, vbTextCompare)
    If Err.Number <> 0 Then
        result = -999
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_property_access() {
        let source = r#"
Sub Test()
    If StrComp(obj.Name, "Test", vbTextCompare) = 0 Then
        MsgBox "Found"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Result: " & StrComp(str1, str2, vbTextCompare)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print StrComp(value1, value2)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_elseif() {
        let source = r#"
Sub Test()
    If StrComp(str1, str2, vbTextCompare) < 0 Then
        result = "Less"
    ElseIf StrComp(str1, str2, vbTextCompare) = 0 Then
        result = "Equal"
    Else
        result = "Greater"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_class_method() {
        let source = r#"
Sub Test()
    Set obj = New StringComparer
    obj.Compare str1, str2
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StringComparer"));
    }

    #[test]
    fn strcomp_binary_search() {
        let source = r#"
Function BinarySearch() As Integer
    compareResult = StrComp(arr(mid), searchValue, vbTextCompare)
    If compareResult = 0 Then
        BinarySearch = mid
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_case_insensitive_equals() {
        let source = r#"
Function EqualsIgnoreCase(s1 As String, s2 As String) As Boolean
    EqualsIgnoreCase = (StrComp(s1, s2, vbTextCompare) = 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }

    #[test]
    fn strcomp_numeric_constant() {
        let source = r#"
Sub Test()
    result = StrComp(str1, str2, 1)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrComp"));
    }
}
