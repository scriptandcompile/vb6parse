
Module documentation in these sub-folders are auto converted into the VB6 library documentation for the github.io webpage. 
They need to follow this specific template:

The first line should have a double ## a single backtick qouting for the function or statement name and then a single empty line:

//! ## `Split` Function
//!

Next, it should have a general description of the function/statements effects followed by a single empty line.

//! Returns a zero-based, one-dimensional array containing a specified number of substrings.
//!

Next is the Syntax section, again preceeded by a double octothorpe, followed by a single empty line, then a code block of type
text that has the syntax for using the function/statement, with a trailing empty line.

//! ## Syntax
//!
//! ```text
//! Split(expression[, delimiter[, limit[, compare]]])
//! ```
//!

The parameters section is next with a double octothorpe header, a single empty line, and then a bulleted list of the arguments
the function/statement takes. These arguments are marked as bold with double asterist quoting, followed by if they are optional
or required and a description of both their type and what those arguments are. If an argument has a specific detail of importance,
such as special values or invalid values, then the line following that argument is indented and marked with a bullet point and the
specific detailed information follows. The section ends with a single empty line.

//! ## Parameters
//!
//! - **expression** (Required): `String` expression containing substrings and delimiters
//! - **delimiter** (Optional): `String` character used to identify substring limits
//!   - If omitted, the space character (" ") is assumed
//! - **limit** (Optional): Number of substrings to be returned; `-1` returns all substrings
//! - **compare** (Optional): Numeric value indicating comparison type (see Compare Settings)
//!

Some parameters have only a small number of set values or values that are predefined as part of the language. In that case,
the follow section will name the argument with a double octothorpe header, have a single empty line, and then list off the
defined names, the values, and what that value represents. Following that is a single empty line. This entire section is optional
and only used when a parameter has need of such a block.

//! ## Compare Settings
//!
//! - `vbBinaryCompare` (0): Perform a binary comparison
//! - `vbTextCompare` (1): Perform a textual comparison
//! - `vbDatabaseCompare` (2): Perform a comparison based on information in your database
//!

The return value section uses a double octothorpe header and lists both the return value and type as well as any special
values that might be returned if the parameter is one of a set of special edge cases. This section is ommited if the
statement has no return value. This section is ended with a single empty line.

//! ## Return Value
//!
//! - Returns a `Variant` containing a one-dimensional array of strings (zero-based)
//! - If `expression` is a zero-length string (""), returns an empty array
//! - If `delimiter` is a zero-length string or not found, returns a single-element array containing the entire expression
//!

The remarks section has any special information which might be pertinent to this function/statement that has not been mentioned
or may have some kind of edge case involving Optional Base, or special handling for strings with zero lengths, nulls, or empty 
arrays. As always, the section is ended with a single empty line.

//! ## Remarks
//!
//! The `Split` function breaks a string into substrings at the specified delimiter and returns them as an array. This is the opposite of the `Join` function, which combines array elements into a single string.
//!
//! - Returns a zero-based array (first element is index 0)
//! - If expression is a zero-length string (""), Split returns an empty array
//! - If delimiter is a zero-length string, a single-element array containing the entire expression is returned
//! - If delimiter is not found, a single-element array containing the entire expression is returned
//! - Delimiter characters are not included in the returned substrings
//! - If `limit` is provided and is less than the number of substrings, the last element contains the remainder of the string (including delimiters)
//! - Multiple consecutive delimiters create empty string elements in the array
//!

'Typical Uses' is another double header section which should contain a bulleted list of common uses for the function/statement, each element of which should have a bulleted common name a colon and a basic explination for that use case. There should be no more than ten items and
the section should end with a single empty line.

//! ## Typical Uses
//!
//! - **Parse CSV Data**: Split comma-separated values
//! - **Extract Words**: Split sentence into individual words
//! - **Process Lines**: Split multiline text into lines
//! - **Parse Paths**: Split file paths into components
//! - **Extract Parameters**: Parse parameter strings
//! - **Data Import**: Process delimited import files
//! - **String Tokenization**: Break strings into tokens
//! * **Configuration Parsing**: Parse config file entries
//!

The 'Common Errors' is an optional section which uses a double # header and should contain a basic description of what can 
cause errors as well as a bolded list of the error numbers, a paranthetical for the common error name, and then when that error
can occur. The error number and common name might be repeated if their are multiple ways to produce the same error number.

If special error handling might be needed or details involving errors an example may be included using a triple # header
title with a code block (marked vb6) that demonstrates the specific needs. As always, a single empty line is used to 
delimit each section.

//! ## Common Errors
//!
//! The Split function itself doesn't typically generate errors with valid inputs, but related operations can:
//!
//! - **Error 13** (Type mismatch): If expression is not a string
//! - **Error 5** (Invalid procedure call): If limit is negative (other than -1)
//! - **Error 9** (Subscript out of range): When accessing array elements beyond bounds
//!
//! ### Always validate inputs and array bounds:
//!
//! ```vb6
//! On Error Resume Next
//! Dim parts() As String
//! parts = Split(text, ",")
//! If Err.Number <> 0 Then
//!     MsgBox "Error splitting text: " & Err.Description
//! End If
//! ```
//!

'Performance Considerations' is an optional double # section which should contain a bulleted list of ways to maximize performance
with the statement/function as well as what to avoid. This section may require an optional triple # named header for an code
example.

//! ## Performance Considerations
//!
//! - Split is very efficient for moderate-sized strings
//! - For very large strings (>1MB), consider processing in chunks
//! - Avoid repeated Split calls in tight loops if possible
//! - Consider caching Split results if reused multiple times
//! - For complex parsing, Split may be slower than manual parsing
//!

The 'Best Practices' section should contain a bulleted list of issues which could arise from using the statement/function
as well as considerations for improving the quality of the resultant code. Each element should have a single bolded statement
followed by clear instructions to apply that tip.

//! ## Best Practices
//!
//! - **Check Array Bounds**: Always verify `UBound` before accessing elements
//! - **Handle Empty Results**: Check if array has elements before processing
//! - **Trim Whitespace**: Use Trim on results to remove unwanted spaces
//! - **Validate Delimiter**: Ensure delimiter is appropriate for data
//! - **Use Limit**: Limit number of splits when only need first few elements
//! - **Handle Edge Cases**: Test with empty strings, missing delimiters
//! - **Consider Alternatives**: For complex parsing, use dedicated parser
//! - **Document Expected Format**: Comment the expected delimited format
//! - **Filter Empty Elements**: Remove empty strings when caused by multiple delimiters
//! - **Combine with Join**: Use Join to reconstruct modified arrays
//!



//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Input | Output |
//! |----------|---------|-------|--------|
//! | Split | String to array | String | Array of strings |
//! | Join | Array to string | Array | String |
//! | Filter | Filter array | Array | Filtered array |
//! | Replace | Replace text | String | String |
//!
//! ## Platform Considerations
//!
//! - Available in VB6, VBA (Office 2000+)
//! - Not available in VBA prior to Office 2000
//! - Returns Variant array (can assign to String array)
//! - Zero-based array (unlike many VB arrays which are 1-based)
//! - Consistent behavior across platforms
//!
//! ## Limitations
//!
//! - Returns zero-based array (may be unexpected in VB6)
//! - Delimiter must be exact match (no regex)
//! - Single delimiter only (can't split on multiple different delimiters)
//! - No built-in trim of results
//! - Empty elements included when multiple consecutive delimiters present
//! - No built-in handling of quoted fields (CSV with commas in quotes)
//! - Maximum array size limited by memory
//!
//! ## Related Functions
//!
//! - `Join`: Combines array elements into a string with delimiter
//! - `Filter`: Returns a subset of array based on filter criteria
//! - `InStr`: Finds position of substring (useful before Split)
//! - `Replace`: Replaces occurrences of substring
//!
//!
//! ## Basic Examples
//!
//! ### Example 1: Split Comma-Separated Values
//!
//! ```vb6
//! Dim text As String
//! Dim parts() As String
//! text = "apple,banana,orange"
//! parts = Split(text, ",")
//! ' parts(0) = "apple"
//! ' parts(1) = "banana"
//! ' parts(2) = "orange"
//! ```
//!
//! ### Example 2: Split Sentence Into Words (Default Space Delimiter)
//!
//! ```vb6
//! Dim sentence As String
//! Dim words() As String
//! sentence = "The quick brown fox"
//! words = Split(sentence)
//! ' words(0) = "The"
//! ' words(1) = "quick"
//! ' words(2) = "brown"
//! ' words(3) = "fox"
//! ```
//!
//! ### Example 3: Split With Limit
//!
//! ```vb6
//! Dim data As String
//! Dim items() As String
//! data = "one,two,three,four,five"
//! items = Split(data, ",", 3)
//! ' items(0) = "one"
//! ' items(1) = "two"
//! ' items(2) = "three,four,five" (remainder)
//! ```
//!
//! ### Example 4: Split Multiline Text
//!
//! ```vb6
//! Dim text As String
//! Dim lines() As String
//! text = "Line 1" & vbCrLf & "Line 2" & vbCrLf & "Line 3"
//! lines = Split(text, vbCrLf)
//! ' lines(0) = "Line 1"
//! ' lines(1) = "Line 2"
//! ' lines(2) = "Line 3"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Parse a CSV line handling quotes
//!
//! ```vb6
//! Function ParseCSVLine(line As String) As String()
//!     ' Simple CSV parsing (doesn't handle quotes)
//!     ParseCSVLine = Split(line, ",")
//! End Function
//! ```
//!
//! ### Pattern 2: Extract Words From Text, Handling Multiple Spaces
//!
//! ```vb6
//! Function GetWords(text As String) As String()
//!     Dim words() As String
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim count As Integer
//!     
//!     words = Split(Trim(text), " ")
//!     
//!     ' Filter out empty strings from multiple spaces
//!     count = 0
//!     For i = LBound(words) To UBound(words)
//!         If Len(words(i)) > 0 Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     ReDim result(0 To count - 1)
//!     count = 0
//!     For i = LBound(words) To UBound(words)
//!         If Len(words(i)) > 0 Then
//!             result(count) = words(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     GetWords = result
//! End Function
//! ```
//!
//! ### Pattern 3: Split File Path Into Components
//!
//! ```vb6
//! Function SplitPath(filePath As String) As String()
//!     Dim delimiter As String
//!     
//!     ' Handle both Windows and Unix paths
//!     If InStr(filePath, "\") > 0 Then
//!         delimiter = "\"
//!     Else
//!         delimiter = "/"
//!     End If
//!     
//!     SplitPath = Split(filePath, delimiter)
//! End Function
//! ```
//!
//! ### Pattern 4: Parse Key=Value Pairs
//!
//! ```vb6
//! Sub ParseKeyValue(kvPair As String, key As String, value As String)
//!     Dim parts() As String
//!     parts = Split(kvPair, "=", 2)
//!     
//!     If UBound(parts) >= 0 Then
//!         key = Trim(parts(0))
//!         If UBound(parts) >= 1 Then
//!             value = Trim(parts(1))
//!         Else
//!             value = ""
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 5: Split Text Into Lines, Handling Different Line Endings
//!
//! ```vb6
//! Function SplitLines(text As String) As String()
//!     Dim normalized As String
//!     
//!     ' Normalize line endings to vbCrLf
//!     normalized = Replace(text, vbCr & vbLf, vbLf)
//!     normalized = Replace(normalized, vbCr, vbLf)
//!     
//!     SplitLines = Split(normalized, vbLf)
//! End Function
//! ```
//!
//! ### Pattern 6: Parse Delimited Data With Custom Delimiter
//!
//! ```vb6
//! Function ParseDelimitedData(data As String, delimiter As String) As Variant
//!     Dim lines() As String
//!     Dim result() As Variant
//!     Dim i As Integer
//!     
//!     lines = Split(data, vbCrLf)
//!     ReDim result(0 To UBound(lines))
//!     
//!     For i = LBound(lines) To UBound(lines)
//!         result(i) = Split(lines(i), delimiter)
//!     Next i
//!     
//!     ParseDelimitedData = result
//! End Function
//! ```
//!
//! ### Pattern 7: Extract Specific Fields From Delimited String
//!
//! ```vb6
//! Function ExtractField(delimitedString As String, _
//!                       delimiter As String, _
//!                       fieldIndex As Integer) As String
//!     Dim fields() As String
//!     fields = Split(delimitedString, delimiter)
//!     
//!     If fieldIndex >= LBound(fields) And fieldIndex <= UBound(fields) Then
//!         ExtractField = fields(fieldIndex)
//!     Else
//!         ExtractField = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: Count Number Of Tokens In String
//!
//! ```vb6
//! Function CountTokens(text As String, delimiter As String) As Integer
//!     Dim tokens() As String
//!     tokens = Split(text, delimiter)
//!     CountTokens = UBound(tokens) - LBound(tokens) + 1
//! End Function
//! ```
//!
//! ### Pattern 9: Split And Reverse The Order
//!
//! ```vb6
//! Function ReverseSplit(text As String, delimiter As String) As String()
//!     Dim parts() As String
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim count As Integer
//!     
//!     parts = Split(text, delimiter)
//!     count = UBound(parts) - LBound(parts)
//!     ReDim result(0 To count)
//!     
//!     For i = 0 To count
//!         result(i) = parts(count - i)
//!     Next i
//!     
//!     ReverseSplit = result
//! End Function
//! ```
//!
//! ### Pattern 10: Split And Remove Empty Elements
//!
//! ```vb6
//! Function SplitNonEmpty(text As String, delimiter As String) As String()
//!     Dim parts() As String
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim count As Integer
//!     
//!     parts = Split(text, delimiter)
//!     
//!     ' Count non-empty elements
//!     count = 0
//!     For i = LBound(parts) To UBound(parts)
//!         If Len(parts(i)) > 0 Then count = count + 1
//!     Next i
//!     
//!     If count = 0 Then
//!         ReDim result(0 To -1)  ' Empty array
//!     Else
//!         ReDim result(0 To count - 1)
//!         count = 0
//!         For i = LBound(parts) To UBound(parts)
//!             If Len(parts(i)) > 0 Then
//!                 result(count) = parts(i)
//!                 count = count + 1
//!             End If
//!         Next i
//!     End If
//!     
//!     SplitNonEmpty = result
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Parse CSV Data With Split
//!
//! ```vb6
//! ' Class: CSVParser
//! Private m_data() As Variant
//! Private m_rowCount As Integer
//! Private m_columnCount As Integer
//!
//! Public Sub LoadCSV(csvText As String, Optional hasHeader As Boolean = True)
//!     Dim lines() As String
//!     Dim i As Integer
//!     Dim startRow As Integer
//!     
//!     ' Split into lines
//!     lines = Split(csvText, vbCrLf)
//!     
//!     If hasHeader Then
//!         startRow = 1
//!         m_rowCount = UBound(lines) - LBound(lines)
//!     Else
//!         startRow = 0
//!         m_rowCount = UBound(lines) - LBound(lines) + 1
//!     End If
//!     
//!     ' Get column count from first data row
//!     If UBound(lines) >= startRow Then
//!         Dim firstRow() As String
//!         firstRow = Split(lines(startRow), ",")
//!         m_columnCount = UBound(firstRow) - LBound(firstRow) + 1
//!     End If
//!     
//!     ' Parse data
//!     ReDim m_data(1 To m_rowCount, 1 To m_columnCount)
//!     
//!     For i = startRow To UBound(lines)
//!         Dim fields() As String
//!         Dim j As Integer
//!         fields = Split(lines(i), ",")
//!         
//!         For j = LBound(fields) To UBound(fields)
//!             If j - LBound(fields) + 1 <= m_columnCount Then
//!                 m_data(i - startRow + 1, j - LBound(fields) + 1) = fields(j)
//!             End If
//!         Next j
//!     Next i
//! End Sub
//!
//! Public Function GetValue(row As Integer, col As Integer) As String
//!     If row >= 1 And row <= m_rowCount And _
//!        col >= 1 And col <= m_columnCount Then
//!         GetValue = m_data(row, col)
//!     Else
//!         GetValue = ""
//!     End If
//! End Function
//!
//! Public Property Get RowCount() As Integer
//!     RowCount = m_rowCount
//! End Property
//!
//! Public Property Get ColumnCount() As Integer
//!     ColumnCount = m_columnCount
//! End Property
//!
//! Public Function GetRow(row As Integer) As Variant
//!     Dim result() As String
//!     Dim i As Integer
//!     
//!     If row >= 1 And row <= m_rowCount Then
//!         ReDim result(1 To m_columnCount)
//!         For i = 1 To m_columnCount
//!             result(i) = m_data(row, i)
//!         Next i
//!         GetRow = result
//!     End If
//! End Function
//! ```
//!
//! ### Example 2: Parse Configuration Files
//!
//! ```vb6
//! ' Module: ConfigFileParser
//! Private m_settings As Object  ' Scripting.Dictionary
//!
//! Public Sub LoadConfig(configText As String)
//!     Dim lines() As String
//!     Dim i As Integer
//!     
//!     Set m_settings = CreateObject("Scripting.Dictionary")
//!     m_settings.CompareMode = vbTextCompare
//!     
//!     lines = Split(configText, vbCrLf)
//!     
//!     For i = LBound(lines) To UBound(lines)
//!         Dim line As String
//!         line = Trim(lines(i))
//!         
//!         ' Skip empty lines and comments
//!         If Len(line) > 0 And Left(line, 1) <> "#" And Left(line, 1) <> ";" Then
//!             Dim parts() As String
//!             parts = Split(line, "=", 2)
//!             
//!             If UBound(parts) >= 1 Then
//!                 Dim key As String
//!                 Dim value As String
//!                 key = Trim(parts(0))
//!                 value = Trim(parts(1))
//!                 
//!                 m_settings(key) = value
//!             End If
//!         End If
//!     Next i
//! End Sub
//!
//! Public Function GetSetting(key As String, Optional defaultValue As String = "") As String
//!     If m_settings.Exists(key) Then
//!         GetSetting = m_settings(key)
//!     Else
//!         GetSetting = defaultValue
//!     End If
//! End Function
//!
//! Public Function GetSettingAsInteger(key As String, Optional defaultValue As Integer = 0) As Integer
//!     If m_settings.Exists(key) Then
//!         If IsNumeric(m_settings(key)) Then
//!             GetSettingAsInteger = CInt(m_settings(key))
//!         Else
//!             GetSettingAsInteger = defaultValue
//!         End If
//!     Else
//!         GetSettingAsInteger = defaultValue
//!     End If
//! End Function
//!
//! Public Function GetSettingList(key As String, delimiter As String) As String()
//!     If m_settings.Exists(key) Then
//!         GetSettingList = Split(m_settings(key), delimiter)
//!     Else
//!         Dim empty() As String
//!         ReDim empty(0 To -1)
//!         GetSettingList = empty
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Process Text With Various Split Operations
//!
//! ```vb6
//! ' Class: TextProcessor
//!
//! Public Function GetParagraphs(text As String) As String()
//!     ' Split by double line breaks
//!     Dim normalized As String
//!     normalized = Replace(text, vbCrLf & vbCrLf, vbLf & vbLf)
//!     normalized = Replace(normalized, vbCr, vbLf)
//!     GetParagraphs = Split(normalized, vbLf & vbLf)
//! End Function
//!
//! Public Function GetSentences(text As String) As String()
//!     Dim temp As String
//!     Dim i As Integer
//!     
//!     ' Simple sentence splitting (doesn't handle abbreviations)
//!     temp = Replace(text, ". ", ".|")
//!     temp = Replace(temp, "! ", "!|")
//!     temp = Replace(temp, "? ", "?|")
//!     
//!     GetSentences = Split(temp, "|")
//! End Function
//!
//! Public Function GetWords(text As String) As String()
//!     Dim cleaned As String
//!     Dim i As Integer
//!     
//!     cleaned = text
//!     ' Remove punctuation
//!     cleaned = Replace(cleaned, ".", " ")
//!     cleaned = Replace(cleaned, ",", " ")
//!     cleaned = Replace(cleaned, "!", " ")
//!     cleaned = Replace(cleaned, "?", " ")
//!     cleaned = Replace(cleaned, ";", " ")
//!     cleaned = Replace(cleaned, ":", " ")
//!     
//!     GetWords = Split(Trim(cleaned), " ")
//! End Function
//!
//! Public Function CountWords(text As String) As Integer
//!     Dim words() As String
//!     Dim count As Integer
//!     Dim i As Integer
//!     
//!     words = GetWords(text)
//!     count = 0
//!     
//!     For i = LBound(words) To UBound(words)
//!         If Len(Trim(words(i))) > 0 Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountWords = count
//! End Function
//!
//! Public Function GetUniqueWords(text As String) As String()
//!     Dim words() As String
//!     Dim dict As Object
//!     Dim i As Integer
//!     Dim result() As String
//!     Dim count As Integer
//!     
//!     Set dict = CreateObject("Scripting.Dictionary")
//!     dict.CompareMode = vbTextCompare
//!     
//!     words = GetWords(text)
//!     
//!     For i = LBound(words) To UBound(words)
//!         Dim word As String
//!         word = Trim(words(i))
//!         If Len(word) > 0 Then
//!             dict(word) = True
//!         End If
//!     Next i
//!     
//!     ReDim result(0 To dict.Count - 1)
//!     Dim keys As Variant
//!     keys = dict.keys
//!     
//!     For i = 0 To dict.Count - 1
//!         result(i) = keys(i)
//!     Next i
//!     
//!     GetUniqueWords = result
//! End Function
//! ```
//!
//! ### Example 4: Import Delimited Data Files
//!
//! ```vb6
//! ' Module: DataImporter
//!
//! Public Function ImportDelimitedFile(filePath As String, _
//!                                     delimiter As String, _
//!                                     Optional hasHeader As Boolean = True) As Variant
//!     Dim fileNum As Integer
//!     Dim fileContent As String
//!     Dim lines() As String
//!     Dim result() As Variant
//!     Dim i As Integer
//!     Dim startRow As Integer
//!     
//!     ' Read file
//!     fileNum = FreeFile
//!     Open filePath For Input As #fileNum
//!     fileContent = Input(LOF(fileNum), #fileNum)
//!     Close #fileNum
//!     
//!     ' Split into lines
//!     lines = Split(fileContent, vbCrLf)
//!     
//!     If hasHeader Then
//!         startRow = 1
//!     Else
//!         startRow = 0
//!     End If
//!     
//!     ' Parse each line
//!     ReDim result(startRow To UBound(lines))
//!     
//!     For i = startRow To UBound(lines)
//!         result(i) = Split(lines(i), delimiter)
//!     Next i
//!     
//!     ImportDelimitedFile = result
//! End Function
//!
//! Public Function GetColumnFromData(data As Variant, columnIndex As Integer) As String()
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim rowCount As Integer
//!     
//!     rowCount = UBound(data) - LBound(data) + 1
//!     ReDim result(0 To rowCount - 1)
//!     
//!     For i = LBound(data) To UBound(data)
//!         Dim row() As String
//!         row = data(i)
//!         
//!         If columnIndex >= LBound(row) And columnIndex <= UBound(row) Then
//!             result(i - LBound(data)) = row(columnIndex)
//!         Else
//!             result(i - LBound(data)) = ""
//!         End If
//!     Next i
//!     
//!     GetColumnFromData = result
//! End Function
//!
//! Public Function FilterRows(data As Variant, columnIndex As Integer, _
//!                            filterValue As String) As Variant
//!     Dim result() As Variant
//!     Dim count As Integer
//!     Dim i As Integer
//!     
//!     ' Count matching rows
//!     count = 0
//!     For i = LBound(data) To UBound(data)
//!         Dim row() As String
//!         row = data(i)
//!         If columnIndex >= LBound(row) And columnIndex <= UBound(row) Then
//!             If row(columnIndex) = filterValue Then
//!                 count = count + 1
//!             End If
//!         End If
//!     Next i
//!     
//!     ' Build result
//!     ReDim result(0 To count - 1)
//!     count = 0
//!     For i = LBound(data) To UBound(data)
//!         row = data(i)
//!         If columnIndex >= LBound(row) And columnIndex <= UBound(row) Then
//!             If row(columnIndex) = filterValue Then
//!                 result(count) = row
//!                 count = count + 1
//!             End If
//!         End If
//!     Next i
//!     
//!     FilterRows = result
//! End Function
//! ```
//!