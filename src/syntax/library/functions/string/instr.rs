//! # `InStr` Function
//!
//! Returns a `Long` specifying the position of the first occurrence of one string within another.
//!
//! ## Syntax
//!
//! ```vb
//! InStr([start, ]string1, string2[, compare])
//! ```
//!
//! ## Parameters
//!
//! - `start` (Optional): Numeric expression that sets the starting position for each search. If omitted, search begins at the first character position. If start contains Null, an error occurs. The start argument is required if compare is specified
//! - `string1` (Required): String expression being searched
//! - `string2` (Required): String expression sought
//! - `compare` (Optional): Specifies the type of string comparison. If compare is `Null`, an error occurs. If compare is omitted, the `Option Compare` setting determines the type of comparison. Specify a valid `LCID` (`LocaleID`) to use locale-specific rules in the comparison
//!
//! ### Compare Parameter Values
//!
//! - `vbUseCompareOption` (-1): Performs a comparison using the setting of the `Option Compare` statement
//! - `vbBinaryCompare` (0): Performs a binary comparison (case-sensitive)
//! - `vbTextCompare` (1): Performs a textual comparison (case-insensitive)
//! - `vbDatabaseCompare` (2): Microsoft Access only. Performs a comparison based on information in your database
//!
//! ## Return Value
//!
//! Returns a `Long`:
//! - If string1 is zero-length: Returns 0
//! - If string1 is `Null`: Returns `Null`
//! - If string2 is zero-length: Returns start
//! - If string2 is `Null`: Returns `Null`
//! - If string2 is not found: Returns 0
//! - If string2 is found within string1: Returns position where match begins
//! - If start > Len(string2): Returns 0
//!
//! ## Remarks
//!
//! The `InStr` function is used for string searching:
//!
//! - Returns the character position of the first occurrence (1-based indexing)
//! - Search is case-sensitive by default (`vbBinaryCompare`)
//! - Use `vbTextCompare` for case-insensitive searching
//! - Start position is 1-based (first character is position 1, not 0)
//! - To find all occurrences, call `InStr` repeatedly with updated start position
//! - `InStrRev` searches from the end of the string backward
//! - The compare parameter affects performance (binary is faster than text)
//! - Commonly used with `Mid`, `Left`, and `Right` functions for string parsing
//! - Returns 0 if substring not found (test with > 0 for found)
//! - `Option Compare` setting affects default comparison when compare is omitted
//!
//! ## Typical Uses
//!
//! 1. **String Searching**: Find if a substring exists in a string
//! 2. **String Parsing**: Locate delimiters for parsing data
//! 3. **Validation**: Check if string contains specific characters
//! 4. **String Extraction**: Find position for `Mid`, `Left`, `Right` operations
//! 5. **Path Manipulation**: Find path separators in file paths
//! 6. **Email Validation**: Locate @ and . in email addresses
//! 7. **Text Processing**: Find keywords or patterns in text
//! 8. **Data Cleanup**: Locate unwanted characters
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple search
//! Dim pos As Long
//! pos = InStr("Hello World", "World")
//! Debug.Print pos  ' Prints: 7
//!
//! ' Example 2: Case-insensitive search
//! Dim pos As Long
//! pos = InStr(1, "Hello World", "world", vbTextCompare)
//! Debug.Print pos  ' Prints: 7
//!
//! ' Example 3: Search from specific position
//! Dim text As String
//! Dim pos As Long
//! text = "apple,banana,apple,orange"
//! pos = InStr(7, text, "apple")
//! Debug.Print pos  ' Prints: 14 (second apple)
//!
//! ' Example 4: Check if substring exists
//! If InStr("user@example.com", "@") > 0 Then
//!     Debug.Print "Valid email format"
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Find all occurrences
//! Sub FindAllOccurrences(text As String, searchText As String)
//!     Dim pos As Long
//!     Dim startPos As Long
//!     
//!     startPos = 1
//!     Do
//!         pos = InStr(startPos, text, searchText)
//!         If pos = 0 Then Exit Do
//!         
//!         Debug.Print "Found at position: " & pos
//!         startPos = pos + Len(searchText)
//!     Loop
//! End Sub
//!
//! ' Pattern 2: Extract substring before delimiter
//! Function GetBeforeDelimiter(text As String, delimiter As String) As String
//!     Dim pos As Long
//!     
//!     pos = InStr(text, delimiter)
//!     If pos > 0 Then
//!         GetBeforeDelimiter = Left$(text, pos - 1)
//!     Else
//!         GetBeforeDelimiter = text
//!     End If
//! End Function
//!
//! ' Pattern 3: Extract substring after delimiter
//! Function GetAfterDelimiter(text As String, delimiter As String) As String
//!     Dim pos As Long
//!     
//!     pos = InStr(text, delimiter)
//!     If pos > 0 Then
//!         GetAfterDelimiter = Mid$(text, pos + Len(delimiter))
//!     Else
//!         GetAfterDelimiter = ""
//!     End If
//! End Function
//!
//! ' Pattern 4: Split string manually
//! Function SplitString(text As String, delimiter As String) As Collection
//!     Dim result As New Collection
//!     Dim pos As Long
//!     Dim startPos As Long
//!     Dim part As String
//!     
//!     startPos = 1
//!     Do
//!         pos = InStr(startPos, text, delimiter)
//!         If pos = 0 Then
//!             ' Add remaining text
//!             result.Add Mid$(text, startPos)
//!             Exit Do
//!         End If
//!         
//!         part = Mid$(text, startPos, pos - startPos)
//!         result.Add part
//!         startPos = pos + Len(delimiter)
//!     Loop
//!     
//!     Set SplitString = result
//! End Function
//!
//! ' Pattern 5: Check for multiple possible substrings
//! Function ContainsAny(text As String, ParamArray searches() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(searches) To UBound(searches)
//!         If InStr(1, text, CStr(searches(i)), vbTextCompare) > 0 Then
//!             ContainsAny = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ContainsAny = False
//! End Function
//!
//! ' Pattern 6: Count occurrences
//! Function CountOccurrences(text As String, searchText As String) As Long
//!     Dim count As Long
//!     Dim pos As Long
//!     Dim startPos As Long
//!     
//!     count = 0
//!     startPos = 1
//!     
//!     Do
//!         pos = InStr(startPos, text, searchText)
//!         If pos = 0 Then Exit Do
//!         
//!         count = count + 1
//!         startPos = pos + Len(searchText)
//!     Loop
//!     
//!     CountOccurrences = count
//! End Function
//!
//! ' Pattern 7: Replace first occurrence
//! Function ReplaceFirst(text As String, findText As String, replaceText As String) As String
//!     Dim pos As Long
//!     
//!     pos = InStr(text, findText)
//!     If pos > 0 Then
//!         ReplaceFirst = Left$(text, pos - 1) & replaceText & _
//!                        Mid$(text, pos + Len(findText))
//!     Else
//!         ReplaceFirst = text
//!     End If
//! End Function
//!
//! ' Pattern 8: Extract file extension
//! Function GetFileExtension(fileName As String) As String
//!     Dim pos As Long
//!     
//!     pos = InStr(fileName, ".")
//!     If pos > 0 Then
//!         ' Find last dot
//!         Do While InStr(pos + 1, fileName, ".") > 0
//!             pos = InStr(pos + 1, fileName, ".")
//!         Loop
//!         GetFileExtension = Mid$(fileName, pos)
//!     Else
//!         GetFileExtension = ""
//!     End If
//! End Function
//!
//! ' Pattern 9: Validate email format
//! Function IsValidEmail(email As String) As Boolean
//!     Dim atPos As Long
//!     Dim dotPos As Long
//!     
//!     email = Trim$(email)
//!     
//!     ' Check for @ symbol
//!     atPos = InStr(email, "@")
//!     If atPos <= 1 Then
//!         IsValidEmail = False
//!         Exit Function
//!     End If
//!     
//!     ' Check for dot after @
//!     dotPos = InStr(atPos + 2, email, ".")
//!     If dotPos = 0 Or dotPos = Len(email) Then
//!         IsValidEmail = False
//!         Exit Function
//!     End If
//!     
//!     IsValidEmail = True
//! End Function
//!
//! ' Pattern 10: Extract domain from URL
//! Function GetDomain(url As String) As String
//!     Dim startPos As Long
//!     Dim endPos As Long
//!     
//!     ' Find start after ://
//!     startPos = InStr(url, "://")
//!     If startPos > 0 Then
//!         startPos = startPos + 3
//!     Else
//!         startPos = 1
//!     End If
//!     
//!     ' Find end at next /
//!     endPos = InStr(startPos, url, "/")
//!     If endPos > 0 Then
//!         GetDomain = Mid$(url, startPos, endPos - startPos)
//!     Else
//!         GetDomain = Mid$(url, startPos)
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Advanced string parser class
//! Public Class StringParser
//!     Private m_text As String
//!     Private m_position As Long
//!     
//!     Public Sub Initialize(text As String)
//!         m_text = text
//!         m_position = 1
//!     End Sub
//!     
//!     Public Function FindNext(searchText As String, Optional caseSensitive As Boolean = True) As Long
//!         Dim pos As Long
//!         Dim compareMode As VbCompareMethod
//!         
//!         compareMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
//!         pos = InStr(m_position, m_text, searchText, compareMode)
//!         
//!         If pos > 0 Then
//!             m_position = pos + Len(searchText)
//!         End If
//!         
//!         FindNext = pos
//!     End Function
//!     
//!     Public Function ExtractBetween(startDelim As String, endDelim As String) As String
//!         Dim startPos As Long
//!         Dim endPos As Long
//!         
//!         startPos = InStr(m_position, m_text, startDelim)
//!         If startPos = 0 Then
//!             ExtractBetween = ""
//!             Exit Function
//!         End If
//!         
//!         startPos = startPos + Len(startDelim)
//!         endPos = InStr(startPos, m_text, endDelim)
//!         
//!         If endPos > 0 Then
//!             ExtractBetween = Mid$(m_text, startPos, endPos - startPos)
//!             m_position = endPos + Len(endDelim)
//!         Else
//!             ExtractBetween = ""
//!         End If
//!     End Function
//!     
//!     Public Sub Reset()
//!         m_position = 1
//!     End Sub
//!     
//!     Public Property Get Position() As Long
//!         Position = m_position
//!     End Property
//!     
//!     Public Property Get RemainingText() As String
//!         RemainingText = Mid$(m_text, m_position)
//!     End Property
//! End Class
//!
//! ' Example 2: Path manipulation utilities
//! Function GetFileName(fullPath As String) As String
//!     Dim pos As Long
//!     Dim lastSlash As Long
//!     
//!     ' Find last backslash or forward slash
//!     lastSlash = 0
//!     pos = 1
//!     
//!     Do
//!         pos = InStr(pos, fullPath, "\")
//!         If pos = 0 Then Exit Do
//!         lastSlash = pos
//!         pos = pos + 1
//!     Loop
//!     
//!     ' Check for forward slash if no backslash found
//!     If lastSlash = 0 Then
//!         pos = 1
//!         Do
//!             pos = InStr(pos, fullPath, "/")
//!             If pos = 0 Then Exit Do
//!             lastSlash = pos
//!             pos = pos + 1
//!         Loop
//!     End If
//!     
//!     If lastSlash > 0 Then
//!         GetFileName = Mid$(fullPath, lastSlash + 1)
//!     Else
//!         GetFileName = fullPath
//!     End If
//! End Function
//!
//! ' Example 3: CSV parser
//! Function ParseCSVLine(csvLine As String) As Collection
//!     Dim result As New Collection
//!     Dim pos As Long
//!     Dim startPos As Long
//!     Dim inQuotes As Boolean
//!     Dim char As String
//!     Dim field As String
//!     Dim i As Long
//!     
//!     startPos = 1
//!     inQuotes = False
//!     
//!     For i = 1 To Len(csvLine)
//!         char = Mid$(csvLine, i, 1)
//!         
//!         If char = """" Then
//!             inQuotes = Not inQuotes
//!         ElseIf char = "," And Not inQuotes Then
//!             field = Mid$(csvLine, startPos, i - startPos)
//!             result.Add Trim$(field)
//!             startPos = i + 1
//!         End If
//!     Next i
//!     
//!     ' Add last field
//!     field = Mid$(csvLine, startPos)
//!     result.Add Trim$(field)
//!     
//!     Set ParseCSVLine = result
//! End Function
//!
//! ' Example 4: Template processor
//! Function ProcessTemplate(template As String, values As Collection) As String
//!     Dim result As String
//!     Dim startPos As Long
//!     Dim endPos As Long
//!     Dim placeholder As String
//!     Dim value As String
//!     Dim i As Long
//!     
//!     result = template
//!     
//!     ' Find all {placeholders}
//!     startPos = 1
//!     Do
//!         startPos = InStr(startPos, result, "{")
//!         If startPos = 0 Then Exit Do
//!         
//!         endPos = InStr(startPos, result, "}")
//!         If endPos = 0 Then Exit Do
//!         
//!         placeholder = Mid$(result, startPos, endPos - startPos + 1)
//!         
//!         ' Look up value (simplified)
//!         For i = 1 To values.Count Step 2
//!             If "{" & values(i) & "}" = placeholder Then
//!                 value = values(i + 1)
//!                 result = Left$(result, startPos - 1) & value & Mid$(result, endPos + 1)
//!                 Exit For
//!             End If
//!         Next i
//!         
//!         startPos = startPos + Len(value)
//!     Loop
//!     
//!     ProcessTemplate = result
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The `InStr` function can raise errors or return Null:
//!
//! - **Type Mismatch (Error 13)**: If arguments are not string-compatible or numeric where expected
//! - **Invalid use of Null (Error 94)**: If string1 or string2 is Null and result is assigned to non-Variant
//! - **Invalid procedure call (Error 5)**: If start < 1
//!
//! ```vb
//! On Error GoTo ErrorHandler
//! Dim pos As Long
//! Dim text As String
//!
//! text = "Hello World"
//! pos = InStr(1, text, "World")
//!
//! If pos > 0 Then
//!     Debug.Print "Found at position: " & pos
//! Else
//!     Debug.Print "Not found"
//! End If
//! Exit Sub
//!
//! ErrorHandler:
//!     MsgBox "Error in InStr: " & Err.Description, vbCritical
//! ```
//!
//! ## Performance Considerations
//!
//! - **Binary vs Text Compare**: Binary comparison (vbBinaryCompare) is faster than text comparison
//! - **String Length**: Performance degrades with very long strings
//! - **Multiple Searches**: For finding all occurrences, performance is linear with number of matches
//! - **Start Position**: Specifying start position avoids re-scanning beginning of string
//! - **Alternative**: For complex pattern matching, consider regular expressions (VBScript.RegExp)
//!
//! ## Best Practices
//!
//! 1. **Test for Found**: Always check if result > 0 before using position
//! 2. **Use vbTextCompare**: For case-insensitive searches, explicitly use vbTextCompare
//! 3. **Start Position**: When searching repeatedly, update start position to avoid infinite loops
//! 4. **Null Handling**: Use Variant for result if strings might be Null
//! 5. **Zero-Length Strings**: Handle empty string cases appropriately
//! 6. **Option Compare**: Be aware of module-level Option Compare setting
//! 7. **Performance**: Use binary compare when case doesn't matter for performance
//!
//! ## Comparison with Other Functions
//!
//! | Function | Purpose | Search Direction |
//! |----------|---------|------------------|
//! | `InStr` | Find substring position | Left to right |
//! | `InStrRev` | Find substring position | Right to left |
//! | `Like` | Pattern matching | N/A |
//! | `StrComp` | Compare strings | N/A |
//! | `Replace` | Find and replace | N/A |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Returns `Long` (32-bit integer), not `Integer`
//! - 1-based indexing (first character is position 1)
//! - Maximum string length is approximately 2GB
//! - Compare parameter affects locale-sensitive comparisons
//! - `Option Compare` statement affects default comparison when compare parameter omitted
//!
//! ## Limitations
//!
//! - Finds only first occurrence (use loop for all occurrences)
//! - No built-in regex or wildcard support
//! - Case-insensitive search (`vbTextCompare`) is slower
//! - Cannot search for multiple substrings in single call
//! - No built-in way to get all positions at once
//! - Performance can degrade with very long strings
//!
//! ## Related Functions
//!
//! - `InStrRev`: Search from end of string backward
//! - `Mid`, `Left`, `Right`: Extract substrings
//! - `Replace`: Find and replace text
//! - `Split`: Split string into array
//! - `Like`: Pattern matching with wildcards
//! - `StrComp`: Compare two strings

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn instr_basic() {
        let source = r#"
Sub Test()
    pos = InStr("Hello World", "World")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_with_start() {
        let source = r#"
Sub Test()
    pos = InStr(5, text, "search")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_with_compare() {
        let source = r#"
Sub Test()
    pos = InStr(1, "Hello", "hello", vbTextCompare)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_if_statement() {
        let source = r#"
Sub Test()
    If InStr(email, "@") > 0 Then
        Debug.Print "Valid"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_in_loop() {
        let source = r"
Sub Test()
    Do While InStr(text, delimiter) > 0
        pos = InStr(text, delimiter)
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_function_return() {
        let source = r#"
Function FindPosition(text As String) As Long
    FindPosition = InStr(text, "target")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_with_mid() {
        let source = r#"
Sub Test()
    pos = InStr(data, ",")
    If pos > 0 Then
        part = Mid$(data, 1, pos - 1)
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_comparison() {
        let source = r#"
Sub Test()
    If InStr(fileName, ".txt") = 0 Then
        MsgBox "Not a text file"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_select_case() {
        let source = r#"
Sub Test()
    Select Case InStr(url, "http")
        Case 1
            Debug.Print "Starts with http"
        Case Is > 0
            Debug.Print "Contains http"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_for_loop() {
        let source = r"
Sub Test()
    Dim i As Long
    For i = 1 To Len(text)
        If InStr(i, text, searchChar) > 0 Then Exit For
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Position: " & InStr(text, "find")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_array_assignment() {
        let source = r"
Sub Test()
    positions(i) = InStr(text, delimiter)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_property_assignment() {
        let source = r"
Sub Test()
    obj.Position = InStr(data, marker)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_delimiterPos = InStr(m_text, m_delimiter)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_with_statement() {
        let source = r"
Sub Test()
    With parser
        .Position = InStr(.Text, .Delimiter)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessPosition(InStr(text, "marker"))
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_concatenation() {
        let source = r#"
Sub Test()
    result = "Found at: " & InStr(text, searchTerm)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_math_expression() {
        let source = r"
Sub Test()
    length = InStr(text, delimiter) - 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_iif() {
        let source = r#"
Sub Test()
    result = IIf(InStr(email, "@") > 0, "Valid", "Invalid")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_multiple_calls() {
        let source = r#"
Sub Test()
    atPos = InStr(email, "@")
    dotPos = InStr(atPos, email, ".")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Position: " & InStr(text, "search")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_collection_add() {
        let source = r"
Sub Test()
    positions.Add InStr(lines(i), delimiter)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_boolean_expression() {
        let source = r#"
Sub Test()
    isValid = InStr(text, "required") > 0 And InStr(text, "approved") > 0
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_nested_call() {
        let source = r#"
Sub Test()
    part = Mid$(text, InStr(text, ":") + 1)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_do_until() {
        let source = r"
Sub Test()
    Do Until InStr(startPos, text, delimiter) = 0
        pos = InStr(startPos, text, delimiter)
        startPos = pos + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_left_function() {
        let source = r#"
Sub Test()
    prefix = Left$(text, InStr(text, " ") - 1)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn instr_parentheses() {
        let source = r"
Sub Test()
    pos = (InStr(text, searchText))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/instr");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
