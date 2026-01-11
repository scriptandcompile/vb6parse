//! # Split Function
//!
//! Returns a zero-based, one-dimensional array containing a specified number of substrings.
//!
//! ## Syntax
//!
//! ```vb
//! Split(expression[, delimiter[, limit[, compare]]])
//! ```
//!
//! ## Parameters
//!
//! - `expression` - Required. String expression containing substrings and delimiters.
//! - `delimiter` - Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter.
//! - `limit` - Optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
//! - `compare` - Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.
//!
//! ## Compare Settings
//!
//! - `vbBinaryCompare` (0): Perform a binary comparison
//! - `vbTextCompare` (1): Perform a textual comparison
//! - `vbDatabaseCompare` (2): Perform a comparison based on information in your database
//!
//! ## Return Value
//!
//! Returns a Variant containing a one-dimensional array of strings. The array is zero-based.
//!
//! ## Remarks
//!
//! The Split function breaks a string into substrings at the specified delimiter and returns them as an array. This is the opposite of the Join function, which combines array elements into a single string.
//!
//! Key characteristics:
//! - Returns a zero-based array (first element is index 0)
//! - If expression is a zero-length string (""), Split returns an empty array
//! - If delimiter is a zero-length string, a single-element array containing the entire expression is returned
//! - If delimiter is not found, a single-element array containing the entire expression is returned
//! - Delimiter characters are not included in the returned substrings
//! - If limit is provided and is less than the number of substrings, the last element contains the remainder of the string (including delimiters)
//! - Multiple consecutive delimiters create empty string elements in the array
//!
//! The Split function is commonly used for:
//! - Parsing delimited data (CSV, TSV, pipe-delimited)
//! - Extracting words from sentences
//! - Processing configuration files
//! - Parsing command-line arguments
//! - Breaking up formatted strings
//! - Converting strings to arrays for processing
//!
//! ## Typical Uses
//!
//! 1. **Parse CSV Data**: Split comma-separated values
//! 2. **Extract Words**: Split sentence into individual words
//! 3. **Process Lines**: Split multiline text into lines
//! 4. **Parse Paths**: Split file paths into components
//! 5. **Extract Parameters**: Parse parameter strings
//! 6. **Data Import**: Process delimited import files
//! 7. **String Tokenization**: Break strings into tokens
//! 8. **Configuration Parsing**: Parse config file entries
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Split comma-separated values
//! Dim text As String
//! Dim parts() As String
//! text = "apple,banana,orange"
//! parts = Split(text, ",")
//! ' parts(0) = "apple"
//! ' parts(1) = "banana"
//! ' parts(2) = "orange"
//! ```
//!
//! ```vb
//! ' Example 2: Split sentence into words (default space delimiter)
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
//! ```vb
//! ' Example 3: Split with limit
//! Dim data As String
//! Dim items() As String
//! data = "one,two,three,four,five"
//! items = Split(data, ",", 3)
//! ' items(0) = "one"
//! ' items(1) = "two"
//! ' items(2) = "three,four,five" (remainder)
//! ```
//!
//! ```vb
//! ' Example 4: Split multiline text
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
//! ### Pattern 1: `ParseCSVLine`
//! Parse a CSV line handling quotes
//! ```vb
//! Function ParseCSVLine(line As String) As String()
//!     ' Simple CSV parsing (doesn't handle quotes)
//!     ParseCSVLine = Split(line, ",")
//! End Function
//! ```
//!
//! ### Pattern 2: `GetWords`
//! Extract words from text, handling multiple spaces
//! ```vb
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
//! ### Pattern 3: `SplitPath`
//! Split file path into components
//! ```vb
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
//! ### Pattern 4: `ParseKeyValue`
//! Parse key=value pairs
//! ```vb
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
//! ### Pattern 5: `SplitLines`
//! Split text into lines, handling different line endings
//! ```vb
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
//! ### Pattern 6: `ParseDelimitedData`
//! Parse delimited data with custom delimiter
//! ```vb
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
//! ### Pattern 7: `ExtractFields`
//! Extract specific fields from delimited string
//! ```vb
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
//! ### Pattern 8: `CountTokens`
//! Count number of tokens in string
//! ```vb
//! Function CountTokens(text As String, delimiter As String) As Integer
//!     Dim tokens() As String
//!     tokens = Split(text, delimiter)
//!     CountTokens = UBound(tokens) - LBound(tokens) + 1
//! End Function
//! ```
//!
//! ### Pattern 9: `ReverseArray`
//! Split and reverse the order
//! ```vb
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
//! ### Pattern 10: `FilterEmptyElements`
//! Split and remove empty elements
//! ```vb
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
//! ### Example 1: `CSVParser` Class
//! Parse CSV data with Split
//! ```vb
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
//! ### Example 2: `ConfigFileParser` Module
//! Parse configuration files
//! ```vb
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
//! ### Example 3: `TextProcessor` Class
//! Process text with various split operations
//! ```vb
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
//! ### Example 4: `DataImporter` Module
//! Import delimited data files
//! ```vb
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
//! ## Error Handling
//!
//! The Split function itself doesn't typically generate errors with valid inputs, but related operations can:
//!
//! - **Error 13** (Type mismatch): If expression is not a string
//! - **Error 5** (Invalid procedure call): If limit is negative (other than -1)
//! - **Error 9** (Subscript out of range): When accessing array elements beyond bounds
//!
//! Always validate inputs and array bounds:
//! ```vb
//! On Error Resume Next
//! Dim parts() As String
//! parts = Split(text, ",")
//! If Err.Number <> 0 Then
//!     MsgBox "Error splitting text: " & Err.Description
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - Split is very efficient for moderate-sized strings
//! - For very large strings (>1MB), consider processing in chunks
//! - Avoid repeated Split calls in tight loops if possible
//! - Consider caching Split results if reused multiple times
//! - For complex parsing, Split may be slower than manual parsing
//!
//! ## Best Practices
//!
//! 1. **Check Array Bounds**: Always verify `UBound` before accessing elements
//! 2. **Handle Empty Results**: Check if array has elements before processing
//! 3. **Trim Whitespace**: Use Trim on results to remove unwanted spaces
//! 4. **Validate Delimiter**: Ensure delimiter is appropriate for data
//! 5. **Use Limit**: Limit number of splits when only need first few elements
//! 6. **Handle Edge Cases**: Test with empty strings, missing delimiters
//! 7. **Consider Alternatives**: For complex parsing, use dedicated parser
//! 8. **Document Expected Format**: Comment the expected delimited format
//! 9. **Filter Empty Elements**: Remove empty strings when caused by multiple delimiters
//! 10. **Combine with Join**: Use Join to reconstruct modified arrays
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
#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn split_basic() {
        let source = r#"
Sub Test()
    Dim parts() As String
    parts = Split("a,b,c", ",")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_with_variable() {
        let source = r#"
Sub Test()
    Dim text As String
    Dim delimiter As String
    Dim result() As String
    text = "one,two,three"
    delimiter = ","
    result = Split(text, delimiter)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_default_delimiter() {
        let source = r#"
Sub Test()
    Dim words() As String
    words = Split("The quick brown fox")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_with_limit() {
        let source = r#"
Sub Test()
    Dim items() As String
    items = Split("a,b,c,d,e", ",", 3)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_if_statement() {
        let source = r#"
Sub Test()
    If UBound(Split(text, ",")) > 0 Then
        MsgBox "Multiple items"
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_function_return() {
        let source = r#"
Function ParseCSV(line As String) As String()
    ParseCSV = Split(line, ",")
End Function
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_variable_assignment() {
        let source = r#"
Sub Test()
    Dim fields() As String
    fields = Split(csvLine, ",")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_array_access() {
        let source = r#"
Sub Test()
    Dim parts() As String
    Dim firstPart As String
    parts = Split("a,b,c", ",")
    firstPart = parts(0)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_in_loop() {
        let source = r#"
Sub Test()
    Dim parts() As String
    Dim i As Integer
    parts = Split(data, ",")
    For i = 0 To UBound(parts)
        Debug.Print parts(i)
    Next i
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_class_usage() {
        let source = r"
Class Parser
    Public Function GetFields(line As String) As String()
        GetFields = Split(line, vbTab)
    End Function
End Class
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_with_statement() {
        let source = r"
Sub Test()
    With parser
        Dim data() As String
        data = Split(.Text, .Delimiter)
    End With
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_elseif() {
        let source = r#"
Sub Test()
    Dim arr() As String
    If delimiter = "," Then
        arr = Split(text, ",")
    ElseIf delimiter = ";" Then
        arr = Split(text, ";")
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_select_case() {
        let source = r#"
Sub Test()
    Dim parts() As String
    Select Case fileType
        Case "CSV"
            parts = Split(line, ",")
        Case "TSV"
            parts = Split(line, vbTab)
    End Select
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_do_while() {
        let source = r#"
Sub Test()
    Do While lineNum <= 10
        Dim fields() As String
        fields = Split(lines(lineNum), ",")
        lineNum = lineNum + 1
    Loop
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_do_until() {
        let source = r#"
Sub Test()
    Do Until EOF(1)
        Dim data() As String
        Line Input #1, currentLine
        data = Split(currentLine, ",")
    Loop
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_while_wend() {
        let source = r#"
Sub Test()
    While i < count
        Dim tokens() As String
        tokens = Split(lines(i), " ")
        i = i + 1
    Wend
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_ubound_check() {
        let source = r#"
Sub Test()
    Dim arr() As String
    arr = Split(text, ",")
    If UBound(arr) >= 0 Then
        MsgBox arr(0)
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_lbound_ubound() {
        let source = r#"
Sub Test()
    Dim parts() As String
    Dim i As Integer
    parts = Split(data, "|")
    For i = LBound(parts) To UBound(parts)
        Debug.Print parts(i)
    Next i
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_vbcrlf() {
        let source = r"
Sub Test()
    Dim lines() As String
    lines = Split(multilineText, vbCrLf)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_nested_function() {
        let source = r#"
Sub Test()
    Dim count As Integer
    count = UBound(Split(text, ",")) + 1
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_join_combination() {
        let source = r#"
Sub Test()
    Dim parts() As String
    Dim result As String
    parts = Split(original, ",")
    result = Join(parts, ";")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_trim_combination() {
        let source = r#"
Sub Test()
    Dim values() As String
    Dim i As Integer
    values = Split(data, ",")
    For i = 0 To UBound(values)
        values(i) = Trim(values(i))
    Next i
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim arr() As String
    arr = Split(text, delimiter)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim fields() As String
    fields = Split(csvData, ",")
    Exit Sub
ErrorHandler:
    MsgBox "Error parsing data"
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_file_path() {
        let source = r#"
Sub Test()
    Dim pathParts() As String
    Dim fileName As String
    pathParts = Split(filePath, "\")
    fileName = pathParts(UBound(pathParts))
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn split_compare_parameter() {
        let source = r"
Sub Test()
    Dim items() As String
    items = Split(text, delimiter, -1, vbTextCompare)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn split_key_value_parsing() {
        let source = r#"
Sub Test()
    Dim kvPair As String
    Dim parts() As String
    Dim key As String
    Dim value As String
    kvPair = "name=John"
    parts = Split(kvPair, "=", 2)
    key = parts(0)
    value = parts(1)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/arrays/split",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
