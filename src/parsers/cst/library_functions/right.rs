//! # Right Function
//!
//! Returns a String containing a specified number of characters from the right side of a string.
//!
//! ## Syntax
//!
//! ```vb
//! Right(string, length)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. String expression from which the rightmost characters are returned.
//! - `length` - Required. Long indicating how many characters to return. If 0, a zero-length string ("") is returned. If greater than or equal to the number of characters in `string`, the entire string is returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the specified number of characters from the right side of the string.
//!
//! ## Remarks
//!
//! The `Right` function extracts a substring from the end of a string. It is commonly used for:
//! - Extracting file extensions
//! - Getting the last N characters of a string
//! - Parsing fixed-width data from the right
//! - Removing prefixes or extracting suffixes
//!
//! **Important Notes**:
//! - If `length` is 0, an empty string ("") is returned
//! - If `length` >= Len(string), the entire string is returned
//! - If `string` is Null, Right returns Null
//! - If `length` is negative, a runtime error occurs (Error 5: Invalid procedure call or argument)
//! - The function is 1-based (counts from position 1, not 0)
//!
//! **Behavior with Different Length Values**:
//!
//! | Length Value | Result |
//! |--------------|--------|
//! | 0 | Empty string ("") |
//! | 1 to Len(string) | Rightmost N characters |
//! | > Len(string) | Entire string |
//! | Negative | Runtime error (Error 5) |
//!
//! ## Typical Uses
//!
//! 1. **File Extensions**: Extract file extension from filename
//! 2. **Data Parsing**: Extract rightmost fields from fixed-width data
//! 3. **String Validation**: Check string endings/suffixes
//! 4. **Number Formatting**: Get last digits of numbers
//! 5. **Path Manipulation**: Extract filename from path
//! 6. **Text Processing**: Remove prefixes, keep suffixes
//! 7. **Data Extraction**: Get trailing characters from codes/IDs
//! 8. **String Truncation**: Keep only the rightmost portion
//!
//! ## Basic Examples
//!
//! ### Example 1: Extract File Extension
//! ```vb
//! Dim filename As String
//! Dim extension As String
//!
//! filename = "document.txt"
//! extension = Right(filename, 4)  ' Returns ".txt"
//! ```
//!
//! ### Example 2: Get Last N Characters
//! ```vb
//! Dim accountNumber As String
//! Dim lastFour As String
//!
//! accountNumber = "1234567890"
//! lastFour = Right(accountNumber, 4)  ' Returns "7890"
//! ```
//!
//! ### Example 3: Check String Ending
//! ```vb
//! Dim filename As String
//!
//! filename = "report.pdf"
//! If Right(filename, 4) = ".pdf" Then
//!     MsgBox "This is a PDF file"
//! End If
//! ```
//!
//! ### Example 4: Extract from Fixed-Width Data
//! ```vb
//! Dim record As String
//! Dim zipCode As String
//!
//! record = "John Doe          New York     10001"
//! zipCode = Right(record, 5)  ' Returns "10001"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `GetFileExtension`
//! ```vb
//! Function GetFileExtension(filename As String) As String
//!     ' Extract file extension including the dot
//!     Dim dotPos As Integer
//!     
//!     dotPos = InStrRev(filename, ".")
//!     
//!     If dotPos > 0 Then
//!         GetFileExtension = Right(filename, Len(filename) - dotPos + 1)
//!     Else
//!         GetFileExtension = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: `GetLastNChars`
//! ```vb
//! Function GetLastNChars(text As String, n As Long) As String
//!     ' Safely get last N characters (won't error if string is too short)
//!     If n <= 0 Then
//!         GetLastNChars = ""
//!     ElseIf n >= Len(text) Then
//!         GetLastNChars = text
//!     Else
//!         GetLastNChars = Right(text, n)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `EndsWithString`
//! ```vb
//! Function EndsWithString(text As String, suffix As String, _
//!                        Optional caseSensitive As Boolean = True) As Boolean
//!     ' Check if string ends with a specific suffix
//!     Dim textEnd As String
//!     Dim suffixLen As Long
//!     
//!     suffixLen = Len(suffix)
//!     
//!     If suffixLen > Len(text) Then
//!         EndsWithString = False
//!         Exit Function
//!     End If
//!     
//!     textEnd = Right(text, suffixLen)
//!     
//!     If caseSensitive Then
//!         EndsWithString = (textEnd = suffix)
//!     Else
//!         EndsWithString = (UCase(textEnd) = UCase(suffix))
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `RemovePrefix`
//! ```vb
//! Function RemovePrefix(text As String, prefixLen As Long) As String
//!     ' Remove prefix by keeping rightmost characters
//!     If prefixLen >= Len(text) Then
//!         RemovePrefix = ""
//!     Else
//!         RemovePrefix = Right(text, Len(text) - prefixLen)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `PadLeftToLength`
//! ```vb
//! Function PadLeftToLength(text As String, totalLength As Long, _
//!                         Optional padChar As String = " ") As String
//!     ' Pad string on the left to reach desired length
//!     Dim currentLen As Long
//!     Dim padding As String
//!     
//!     currentLen = Len(text)
//!     
//!     If currentLen >= totalLength Then
//!         PadLeftToLength = Right(text, totalLength)
//!     Else
//!         padding = String(totalLength - currentLen, padChar)
//!         PadLeftToLength = padding & text
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 6: `GetFilenameFromPath`
//! ```vb
//! Function GetFilenameFromPath(fullPath As String) As String
//!     ' Extract filename from full path
//!     Dim slashPos As Integer
//!     
//!     slashPos = InStrRev(fullPath, "\")
//!     
//!     If slashPos > 0 Then
//!         GetFilenameFromPath = Right(fullPath, Len(fullPath) - slashPos)
//!     Else
//!         GetFilenameFromPath = fullPath
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: `TruncateLeft`
//! ```vb
//! Function TruncateLeft(text As String, maxLength As Long, _
//!                      Optional ellipsis As String = "...") As String
//!     ' Truncate from left, keeping rightmost characters
//!     Dim textLen As Long
//!     
//!     textLen = Len(text)
//!     
//!     If textLen <= maxLength Then
//!         TruncateLeft = text
//!     Else
//!         TruncateLeft = ellipsis & Right(text, maxLength - Len(ellipsis))
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: `FormatAccountNumber`
//! ```vb
//! Function FormatAccountNumber(accountNum As String) As String
//!     ' Format account number showing only last 4 digits
//!     Dim lastFour As String
//!     
//!     If Len(accountNum) > 4 Then
//!         lastFour = Right(accountNum, 4)
//!         FormatAccountNumber = "****" & lastFour
//!     Else
//!         FormatAccountNumber = accountNum
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `ExtractDomainExtension`
//! ```vb
//! Function ExtractDomainExtension(url As String) As String
//!     ' Extract domain extension (.com, .org, etc.)
//!     Dim dotPos As Integer
//!     
//!     ' Find last dot
//!     dotPos = InStrRev(url, ".")
//!     
//!     If dotPos > 0 Then
//!         ExtractDomainExtension = Right(url, Len(url) - dotPos + 1)
//!     Else
//!         ExtractDomainExtension = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `GetTrailingDigits`
//! ```vb
//! Function GetTrailingDigits(text As String) As String
//!     ' Extract trailing numeric characters
//!     Dim i As Integer
//!     Dim digitCount As Integer
//!     
//!     digitCount = 0
//!     
//!     For i = Len(text) To 1 Step -1
//!         If IsNumeric(Mid(text, i, 1)) Then
//!             digitCount = digitCount + 1
//!         Else
//!             Exit For
//!         End If
//!     Next i
//!     
//!     If digitCount > 0 Then
//!         GetTrailingDigits = Right(text, digitCount)
//!     Else
//!         GetTrailingDigits = ""
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: File Extension Validator
//! ```vb
//! ' Comprehensive file extension validation and management
//! Class FileExtensionValidator
//!     Private m_validExtensions As Collection
//!     
//!     Public Sub Initialize()
//!         Set m_validExtensions = New Collection
//!     End Sub
//!     
//!     Public Sub AddValidExtension(extension As String)
//!         Dim ext As String
//!         
//!         ' Normalize extension (add dot if missing)
//!         If Left(extension, 1) <> "." Then
//!             ext = "." & extension
//!         Else
//!             ext = extension
//!         End If
//!         
//!         On Error Resume Next
//!         m_validExtensions.Add ext, UCase(ext)
//!         On Error GoTo 0
//!     End Sub
//!     
//!     Public Function IsValidFile(filename As String) As Boolean
//!         Dim ext As String
//!         Dim dotPos As Integer
//!         
//!         dotPos = InStrRev(filename, ".")
//!         
//!         If dotPos = 0 Then
//!             IsValidFile = False
//!             Exit Function
//!         End If
//!         
//!         ext = Right(filename, Len(filename) - dotPos + 1)
//!         
//!         On Error Resume Next
//!         Dim temp As String
//!         temp = m_validExtensions(UCase(ext))
//!         IsValidFile = (Err.Number = 0)
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function GetExtension(filename As String) As String
//!         Dim dotPos As Integer
//!         
//!         dotPos = InStrRev(filename, ".")
//!         
//!         If dotPos > 0 Then
//!             GetExtension = Right(filename, Len(filename) - dotPos + 1)
//!         Else
//!             GetExtension = ""
//!         End If
//!     End Function
//!     
//!     Public Function ChangeExtension(filename As String, newExt As String) As String
//!         Dim dotPos As Integer
//!         Dim baseName As String
//!         Dim ext As String
//!         
//!         dotPos = InStrRev(filename, ".")
//!         
//!         If dotPos > 0 Then
//!             baseName = Left(filename, dotPos - 1)
//!         Else
//!             baseName = filename
//!         End If
//!         
//!         ' Normalize new extension
//!         If Left(newExt, 1) <> "." Then
//!             ext = "." & newExt
//!         Else
//!             ext = newExt
//!         End If
//!         
//!         ChangeExtension = baseName & ext
//!     End Function
//!     
//!     Public Function GetValidExtensions() As String
//!         Dim ext As Variant
//!         Dim result As String
//!         
//!         result = ""
//!         
//!         For Each ext In m_validExtensions
//!             If result <> "" Then result = result & ", "
//!             result = result & ext
//!         Next ext
//!         
//!         GetValidExtensions = result
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Path Parser Module
//! ```vb
//! ' Parse and manipulate file paths
//! Module PathParser
//!     Public Function GetFilename(fullPath As String) As String
//!         ' Extract filename from full path
//!         Dim slashPos As Integer
//!         
//!         slashPos = InStrRev(fullPath, "\")
//!         
//!         If slashPos > 0 Then
//!             GetFilename = Right(fullPath, Len(fullPath) - slashPos)
//!         Else
//!             GetFilename = fullPath
//!         End If
//!     End Function
//!     
//!     Public Function GetDirectory(fullPath As String) As String
//!         ' Extract directory from full path
//!         Dim slashPos As Integer
//!         
//!         slashPos = InStrRev(fullPath, "\")
//!         
//!         If slashPos > 0 Then
//!             GetDirectory = Left(fullPath, slashPos - 1)
//!         Else
//!             GetDirectory = ""
//!         End If
//!     End Function
//!     
//!     Public Function GetFilenameWithoutExtension(fullPath As String) As String
//!         ' Get filename without extension
//!         Dim filename As String
//!         Dim dotPos As Integer
//!         
//!         filename = GetFilename(fullPath)
//!         dotPos = InStrRev(filename, ".")
//!         
//!         If dotPos > 0 Then
//!             GetFilenameWithoutExtension = Left(filename, dotPos - 1)
//!         Else
//!             GetFilenameWithoutExtension = filename
//!         End If
//!     End Function
//!     
//!     Public Function GetExtension(fullPath As String) As String
//!         ' Get file extension including dot
//!         Dim filename As String
//!         Dim dotPos As Integer
//!         
//!         filename = GetFilename(fullPath)
//!         dotPos = InStrRev(filename, ".")
//!         
//!         If dotPos > 0 Then
//!             GetExtension = Right(filename, Len(filename) - dotPos + 1)
//!         Else
//!             GetExtension = ""
//!         End If
//!     End Function
//!     
//!     Public Function CombinePath(directory As String, filename As String) As String
//!         ' Combine directory and filename
//!         If Right(directory, 1) = "\" Then
//!             CombinePath = directory & filename
//!         Else
//!             CombinePath = directory & "\" & filename
//!         End If
//!     End Function
//!     
//!     Public Function GetParentDirectory(fullPath As String) As String
//!         ' Get parent directory
//!         Dim dir As String
//!         Dim slashPos As Integer
//!         
//!         dir = GetDirectory(fullPath)
//!         slashPos = InStrRev(dir, "\")
//!         
//!         If slashPos > 0 Then
//!             GetParentDirectory = Left(dir, slashPos - 1)
//!         Else
//!             GetParentDirectory = ""
//!         End If
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: String Suffix Matcher
//! ```vb
//! ' Match and validate string suffixes
//! Class StringSuffixMatcher
//!     Private m_caseSensitive As Boolean
//!     
//!     Public Sub Initialize(Optional caseSensitive As Boolean = True)
//!         m_caseSensitive = caseSensitive
//!     End Sub
//!     
//!     Public Function EndsWith(text As String, suffix As String) As Boolean
//!         ' Check if string ends with suffix
//!         Dim textEnd As String
//!         Dim suffixLen As Long
//!         
//!         suffixLen = Len(suffix)
//!         
//!         If suffixLen > Len(text) Then
//!             EndsWith = False
//!             Exit Function
//!         End If
//!         
//!         textEnd = Right(text, suffixLen)
//!         
//!         If m_caseSensitive Then
//!             EndsWith = (textEnd = suffix)
//!         Else
//!             EndsWith = (UCase(textEnd) = UCase(suffix))
//!         End If
//!     End Function
//!     
//!     Public Function EndsWithAny(text As String, suffixes() As String) As Boolean
//!         ' Check if string ends with any of the provided suffixes
//!         Dim i As Integer
//!         
//!         For i = LBound(suffixes) To UBound(suffixes)
//!             If EndsWith(text, suffixes(i)) Then
//!                 EndsWithAny = True
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         EndsWithAny = False
//!     End Function
//!     
//!     Public Function RemoveSuffix(text As String, suffix As String) As String
//!         ' Remove suffix if present
//!         If EndsWith(text, suffix) Then
//!             RemoveSuffix = Left(text, Len(text) - Len(suffix))
//!         Else
//!             RemoveSuffix = text
//!         End If
//!     End Function
//!     
//!     Public Function GetMatchingSuffix(text As String, suffixes() As String) As String
//!         ' Return the matching suffix, or empty string
//!         Dim i As Integer
//!         
//!         For i = LBound(suffixes) To UBound(suffixes)
//!             If EndsWith(text, suffixes(i)) Then
//!                 GetMatchingSuffix = suffixes(i)
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         GetMatchingSuffix = ""
//!     End Function
//!     
//!     Public Function ReplaceSuffix(text As String, oldSuffix As String, _
//!                                  newSuffix As String) As String
//!         ' Replace suffix if present
//!         If EndsWith(text, oldSuffix) Then
//!             ReplaceSuffix = Left(text, Len(text) - Len(oldSuffix)) & newSuffix
//!         Else
//!             ReplaceSuffix = text
//!         End If
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Account Number Formatter
//! ```vb
//! ' Format and mask account numbers securely
//! Class AccountNumberFormatter
//!     Private m_maskChar As String
//!     Private m_visibleDigits As Integer
//!     
//!     Public Sub Initialize(Optional maskChar As String = "*", _
//!                          Optional visibleDigits As Integer = 4)
//!         m_maskChar = maskChar
//!         m_visibleDigits = visibleDigits
//!     End Sub
//!     
//!     Public Function FormatAccountNumber(accountNum As String) As String
//!         ' Format account number showing only last N digits
//!         Dim lastDigits As String
//!         Dim maskedLength As Integer
//!         
//!         If Len(accountNum) <= m_visibleDigits Then
//!             FormatAccountNumber = accountNum
//!             Exit Function
//!         End If
//!         
//!         lastDigits = Right(accountNum, m_visibleDigits)
//!         maskedLength = Len(accountNum) - m_visibleDigits
//!         
//!         FormatAccountNumber = String(maskedLength, m_maskChar) & lastDigits
//!     End Function
//!     
//!     Public Function FormatWithSpaces(accountNum As String, _
//!                                     Optional groupSize As Integer = 4) As String
//!         ' Format with spaces every N characters
//!         Dim formatted As String
//!         Dim i As Integer
//!         Dim maskedNum As String
//!         
//!         maskedNum = FormatAccountNumber(accountNum)
//!         formatted = ""
//!         
//!         For i = 1 To Len(maskedNum) Step groupSize
//!             If i > 1 Then formatted = formatted & " "
//!             formatted = formatted & Mid(maskedNum, i, groupSize)
//!         Next i
//!         
//!         FormatWithSpaces = formatted
//!     End Function
//!     
//!     Public Function GetLastDigits(accountNum As String) As String
//!         ' Get only the visible digits
//!         If Len(accountNum) <= m_visibleDigits Then
//!             GetLastDigits = accountNum
//!         Else
//!             GetLastDigits = Right(accountNum, m_visibleDigits)
//!         End If
//!     End Function
//!     
//!     Public Function ValidateAccountNumber(accountNum As String, _
//!                                          minLength As Integer, _
//!                                          maxLength As Integer) As Boolean
//!         ' Validate account number length
//!         Dim numLen As Integer
//!         
//!         numLen = Len(accountNum)
//!         ValidateAccountNumber = (numLen >= minLength And numLen <= maxLength)
//!     End Function
//!     
//!     Public Sub SetMaskChar(maskChar As String)
//!         m_maskChar = maskChar
//!     End Sub
//!     
//!     Public Sub SetVisibleDigits(visibleDigits As Integer)
//!         m_visibleDigits = visibleDigits
//!     End Sub
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `Right` function generates runtime errors in specific situations:
//!
//! **Error 5: Invalid procedure call or argument**
//! - Occurs when `length` parameter is negative
//!
//! **Error 94: Invalid use of Null**
//! - Occurs when `string` parameter is Null
//!
//! Example error handling:
//!
//! ```vb
//! On Error Resume Next
//! result = Right(userInput, charCount)
//! If Err.Number <> 0 Then
//!     MsgBox "Error extracting characters: " & Err.Description
//!     result = ""
//! End If
//! On Error GoTo 0
//! ```
//!
//! Best practice for safe usage:
//!
//! ```vb
//! Function SafeRight(text As String, length As Long) As String
//!     If IsNull(text) Then
//!         SafeRight = ""
//!         Exit Function
//!     End If
//!     
//!     If length < 0 Then
//!         SafeRight = ""
//!         Exit Function
//!     End If
//!     
//!     If length = 0 Then
//!         SafeRight = ""
//!     ElseIf length >= Len(text) Then
//!         SafeRight = text
//!     Else
//!         SafeRight = Right(text, length)
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Right` is a very fast string function
//! - Performance is O(n) where n is the length parameter
//! - More efficient than using Mid to extract from the end
//! - For large strings, consider caching Len(string) if called multiple times
//! - No significant performance difference between Right and string slicing
//!
//! ## Best Practices
//!
//! 1. **Validate Length**: Check that length parameter is valid before calling
//! 2. **Handle Null**: Check for Null strings if data source is uncertain
//! 3. **Use with `InStrRev`**: Combine with `InStrRev` for finding from right
//! 4. **Document Intent**: Make clear why extracting from right vs left
//! 5. **Consider Edge Cases**: Handle empty strings, length = 0, length > string length
//! 6. **Use for File Extensions**: Preferred method for extracting file extensions
//! 7. **Combine with Trim**: Often useful to Trim before using Right
//! 8. **Cache String Length**: If using Len(string) multiple times, cache it
//! 9. **Avoid Magic Numbers**: Use named constants for length values
//! 10. **Test Boundary Conditions**: Test with length = 0, 1, Len(string), Len(string)+1
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Parameters | Use Case |
//! |----------|---------|------------|----------|
//! | **Right** | Extract from right | (string, length) | Get last N characters, file extensions |
//! | **Left** | Extract from left | (string, length) | Get first N characters, prefixes |
//! | **Mid** | Extract from middle | (string, start, [length]) | Get substring from any position |
//! | **`InStrRev`** | Find from right | (string, substring) | Find position searching from right |
//! | **Len** | String length | (string) | Get total length |
//! | **`RTrim`** | Remove right spaces | (string) | Remove trailing whitespace |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VB6 and VBA
//! - Behavior consistent across all platforms
//! - In VB.NET, replaced by String.Substring or string slicing
//! - `RightB` and `RightB`$ variants exist for byte data
//! - Right$ returns String type, Right returns Variant
//!
//! ## Limitations
//!
//! - Cannot use negative length to count from a different position
//! - No built-in way to extract "all but first N" characters (use Len arithmetic)
//! - Raises error on negative length instead of returning empty string
//! - No case-insensitive comparison built-in
//! - No Unicode-aware variant (uses ANSI/DBCS)
//!
//! ## Related Functions
//!
//! - `Left`: Returns characters from the left side of a string
//! - `Mid`: Returns characters from any position in a string
//! - `InStrRev`: Finds position of substring searching from right to left
//! - `Len`: Returns the length of a string
//! - `RTrim`: Removes trailing spaces from a string
//! - `InStr`: Finds position of substring searching from left to right

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn right_basic() {
        let source = r#"
Dim result As String
result = Right("Hello World", 5)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_file_extension() {
        let source = r#"
Dim extension As String
extension = Right(filename, 4)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_if_statement() {
        let source = r#"
If Right(filename, 4) = ".txt" Then
    MsgBox "Text file"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_function_return() {
        let source = r#"
Function GetLastFour(s As String) As String
    GetLastFour = Right(s, 4)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_variable_assignment() {
        let source = r#"
Dim lastChars As String
lastChars = Right(inputText, charCount)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_msgbox() {
        let source = r#"
MsgBox "Last 3 chars: " & Right(text, 3)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_debug_print() {
        let source = r#"
Debug.Print Right("Testing", 4)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_select_case() {
        let source = r#"
Select Case Right(filename, 4)
    Case ".txt"
        ProcessText
    Case ".doc"
        ProcessDoc
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_class_usage() {
        let source = r#"
Private m_suffix As String

Public Sub ExtractSuffix()
    m_suffix = Right(m_text, 10)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_with_statement() {
        let source = r#"
With TextBox1
    .Text = Right(.Text, 20)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_elseif() {
        let source = r#"
If Right(s, 4) = ".exe" Then
    fileType = "Executable"
ElseIf Right(s, 4) = ".dll" Then
    fileType = "Library"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_for_loop() {
        let source = r#"
For i = 1 To 10
    parts(i) = Right(lines(i), 5)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_do_while() {
        let source = r#"
Do While Right(buffer, 2) <> vbCrLf
    buffer = buffer & ReadChar()
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_do_until() {
        let source = r#"
Do Until Right(data, 1) = ";"
    data = data & GetNextByte()
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_while_wend() {
        let source = r#"
While Right(line, 1) = " "
    line = Left(line, Len(line) - 1)
Wend
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_parentheses() {
        let source = r#"
Dim val As String
val = (Right(input, 10))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_iif() {
        let source = r#"
Dim ext As String
ext = IIf(hasExtension, Right(name, 4), ".txt")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_nested() {
        let source = r#"
Dim result As String
result = Right(Right(fullPath, 20), 10)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_array_assignment() {
        let source = r#"
Dim suffixes(10) As String
suffixes(i) = Right(words(i), 3)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_property_assignment() {
        let source = r#"
Set obj = New StringHelper
obj.Suffix = Right(fullString, 15)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_function_argument() {
        let source = r#"
Call ProcessExtension(Right(filename, 4))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_concatenation() {
        let source = r#"
Dim msg As String
msg = "Extension: " & Right(file, 3)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_comparison() {
        let source = r#"
If Right(url, 4) = ".com" Or Right(url, 4) = ".net" Then
    ValidDomain = True
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_with_len() {
        let source = r#"
Dim remaining As String
remaining = Right(text, Len(text) - 5)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("remaining"));
    }

    #[test]
    fn right_trim_combination() {
        let source = r#"
Dim cleaned As String
cleaned = Right(Trim(input), 10)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_error_handling() {
        let source = r#"
On Error Resume Next
result = Right(userInput, count)
If Err.Number <> 0 Then
    result = ""
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn right_on_error_goto() {
        let source = r#"
Sub ExtractSuffix()
    On Error GoTo ErrorHandler
    Dim suffix As String
    suffix = Right(text, n)
    Exit Sub
ErrorHandler:
    MsgBox "Error extracting suffix"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Right"));
        assert!(text.contains("Identifier"));
    }
}
