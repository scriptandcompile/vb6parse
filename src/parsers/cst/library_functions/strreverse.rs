//! VB6 `StrReverse` Function
//!
//! The `StrReverse` function returns a string in which the character order of a specified string is reversed.
//!
//! ## Syntax
//! ```vb6
//! StrReverse(expression)
//! ```
//!
//! ## Parameters
//! - `expression`: Required. String expression whose characters are to be reversed. If `expression` is a zero-length string (""), a zero-length string is returned.
//!
//! ## Returns
//! Returns a `String` with the characters in reverse order.
//!
//! ## Remarks
//! The `StrReverse` function reverses the order of characters in a string:
//!
//! - **Character-by-character reversal**: Reverses individual characters, not words
//! - **Unicode support**: Works correctly with Unicode characters
//! - **Empty string handling**: Returns empty string if input is empty
//! - **Null handling**: Returns `Null` if `expression` is `Null`
//! - **Preserves spaces**: Spaces are treated like any other character and reversed
//! - **Case preserved**: Original case of characters is maintained
//! - **Single pass**: Efficient single-pass algorithm
//! - **VB6/VBA only**: Available in VB6 and VBA, not in `VBScript`
//!
//! ### Common Use Cases
//! - Reversing strings for display or analysis
//! - Palindrome checking (compare string with its reverse)
//! - Text transformations and puzzles
//! - Data obfuscation (simple, not secure)
//! - Mirror text effects
//! - String manipulation algorithms
//!
//! ### Comparison with Manual Reversal
//! `StrReverse` is more efficient than manually reversing with loops:
//! ```vb6
//! ' Using StrReverse (preferred)
//! reversed = StrReverse(original)
//!
//! ' Manual reversal (slower)
//! For i = Len(original) To 1 Step -1
//!     reversed = reversed & Mid$(original, i, 1)
//! Next i
//! ```
//!
//! ## Typical Uses
//! 1. **Palindrome Detection**: Check if a string reads the same forwards and backwards
//! 2. **Text Effects**: Create mirror or reversed text displays
//! 3. **String Analysis**: Analyze patterns in reversed strings
//! 4. **Data Transformation**: Transform data for specific algorithms
//! 5. **Puzzles and Games**: Implement word games and puzzles
//! 6. **File Processing**: Process files that store data in reverse order
//! 7. **Encoding**: Simple (non-cryptographic) string obfuscation
//! 8. **Testing**: Generate test data with predictable patterns
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic String Reversal
//! ```vb6
//! Dim original As String
//! Dim reversed As String
//!
//! original = "Hello"
//! reversed = StrReverse(original)  ' "olleH"
//!
//! original = "VB6"
//! reversed = StrReverse(original)  ' "6BV"
//!
//! original = "12345"
//! reversed = StrReverse(original)  ' "54321"
//! ```
//!
//! ### Example 2: Palindrome Check
//! ```vb6
//! Function IsPalindrome(text As String) As Boolean
//!     Dim normalized As String
//!     
//!     ' Remove spaces and convert to lowercase for comparison
//!     normalized = Replace(LCase$(text), " ", "")
//!     
//!     ' Compare with reversed version
//!     IsPalindrome = (normalized = StrReverse(normalized))
//! End Function
//!
//! ' Examples:
//! ' IsPalindrome("racecar") = True
//! ' IsPalindrome("A man a plan a canal Panama") = True (after normalization)
//! ' IsPalindrome("hello") = False
//! ```
//!
//! ### Example 3: Reverse Words in Sentence
//! ```vb6
//! Function ReverseWords(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     words = Split(sentence, " ")
//!     
//!     result = ""
//!     For i = UBound(words) To LBound(words) Step -1
//!         If i < UBound(words) Then result = result & " "
//!         result = result & words(i)
//!     Next i
//!     
//!     ReverseWords = result
//! End Function
//!
//! ' Example: "Hello World" becomes "World Hello"
//! ```
//!
//! ### Example 4: Simple Obfuscation
//! ```vb6
//! Function ObfuscateString(text As String) As String
//!     ' Simple, non-secure obfuscation
//!     ObfuscateString = StrReverse(text)
//! End Function
//!
//! Function DeobfuscateString(text As String) As String
//!     ' Reverse the obfuscation
//!     DeobfuscateString = StrReverse(text)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Check Palindrome (Case-Insensitive)
//! ```vb6
//! Function IsPalindromeIgnoreCase(text As String) As Boolean
//!     Dim lower As String
//!     lower = LCase$(text)
//!     IsPalindromeIgnoreCase = (lower = StrReverse(lower))
//! End Function
//! ```
//!
//! ### Pattern 2: Reverse Each Word
//! ```vb6
//! Function ReverseEachWord(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     
//!     words = Split(sentence, " ")
//!     For i = LBound(words) To UBound(words)
//!         words(i) = StrReverse(words(i))
//!     Next i
//!     
//!     ReverseEachWord = Join(words, " ")
//! End Function
//! ```
//!
//! ### Pattern 3: Get Last N Characters Efficiently
//! ```vb6
//! Function GetLastNChars(text As String, n As Integer) As String
//!     Dim reversed As String
//!     reversed = StrReverse(text)
//!     GetLastNChars = StrReverse(Left$(reversed, n))
//! End Function
//! ```
//!
//! ### Pattern 4: Check If Strings Are Reverses
//! ```vb6
//! Function AreReverses(str1 As String, str2 As String) As Boolean
//!     AreReverses = (str1 = StrReverse(str2))
//! End Function
//! ```
//!
//! ### Pattern 5: Reverse File Extension
//! ```vb6
//! Function ReverseExtension(filename As String) As String
//!     Dim dotPos As Integer
//!     Dim name As String
//!     Dim ext As String
//!     
//!     dotPos = InStrRev(filename, ".")
//!     If dotPos > 0 Then
//!         name = Left$(filename, dotPos - 1)
//!         ext = Mid$(filename, dotPos + 1)
//!         ReverseExtension = name & "." & StrReverse(ext)
//!     Else
//!         ReverseExtension = filename
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 6: Mirror Text Display
//! ```vb6
//! Function CreateMirrorText(text As String) As String
//!     CreateMirrorText = text & " | " & StrReverse(text)
//! End Function
//! ```
//!
//! ### Pattern 7: Reverse and Uppercase
//! ```vb6
//! Function ReverseAndUpper(text As String) As String
//!     ReverseAndUpper = UCase$(StrReverse(text))
//! End Function
//! ```
//!
//! ### Pattern 8: Find Reverse Match in Array
//! ```vb6
//! Function FindReverseMatch(arr() As String, searchValue As String) As Integer
//!     Dim i As Integer
//!     Dim reversed As String
//!     
//!     reversed = StrReverse(searchValue)
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) = reversed Then
//!             FindReverseMatch = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     FindReverseMatch = -1
//! End Function
//! ```
//!
//! ### Pattern 9: Reverse Between Delimiters
//! ```vb6
//! Function ReverseBetween(text As String, startDelim As String, endDelim As String) As String
//!     Dim startPos As Integer
//!     Dim endPos As Integer
//!     Dim middle As String
//!     
//!     startPos = InStr(text, startDelim)
//!     endPos = InStr(startPos + Len(startDelim), text, endDelim)
//!     
//!     If startPos > 0 And endPos > startPos Then
//!         middle = Mid$(text, startPos + Len(startDelim), endPos - startPos - Len(startDelim))
//!         ReverseBetween = Left$(text, startPos + Len(startDelim) - 1) & _
//!                         StrReverse(middle) & _
//!                         Mid$(text, endPos)
//!     Else
//!         ReverseBetween = text
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: Alternate Characters Reversed
//! ```vb6
//! Function AlternateReverse(text As String) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim reversed As String
//!     
//!     reversed = StrReverse(text)
//!     result = ""
//!     
//!     For i = 1 To Len(text)
//!         If i Mod 2 = 1 Then
//!             result = result & Mid$(text, i, 1)
//!         Else
//!             result = result & Mid$(reversed, i, 1)
//!         End If
//!     Next i
//!     
//!     AlternateReverse = result
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Palindrome Checker Class
//! ```vb6
//! ' Class: PalindromeChecker
//! ' Checks various types of palindromes
//! Option Explicit
//!
//! Private m_IgnoreCase As Boolean
//! Private m_IgnoreSpaces As Boolean
//! Private m_IgnorePunctuation As Boolean
//!
//! Public Sub Initialize(Optional ignoreCase As Boolean = True, _
//!                       Optional ignoreSpaces As Boolean = True, _
//!                       Optional ignorePunctuation As Boolean = False)
//!     m_IgnoreCase = ignoreCase
//!     m_IgnoreSpaces = ignoreSpaces
//!     m_IgnorePunctuation = ignorePunctuation
//! End Sub
//!
//! Public Function IsPalindrome(text As String) As Boolean
//!     Dim normalized As String
//!     normalized = NormalizeText(text)
//!     IsPalindrome = (normalized = StrReverse(normalized))
//! End Function
//!
//! Public Function GetLongestPalindromeSubstring(text As String) As String
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim substring As String
//!     Dim longest As String
//!     
//!     longest = ""
//!     
//!     For i = 1 To Len(text)
//!         For j = i To Len(text)
//!             substring = Mid$(text, i, j - i + 1)
//!             If IsPalindrome(substring) And Len(substring) > Len(longest) Then
//!                 longest = substring
//!             End If
//!         Next j
//!     Next i
//!     
//!     GetLongestPalindromeSubstring = longest
//! End Function
//!
//! Private Function NormalizeText(text As String) As String
//!     Dim result As String
//!     result = text
//!     
//!     If m_IgnoreCase Then
//!         result = LCase$(result)
//!     End If
//!     
//!     If m_IgnoreSpaces Then
//!         result = Replace(result, " ", "")
//!     End If
//!     
//!     If m_IgnorePunctuation Then
//!         result = RemovePunctuation(result)
//!     End If
//!     
//!     NormalizeText = result
//! End Function
//!
//! Private Function RemovePunctuation(text As String) As String
//!     Dim i As Integer
//!     Dim char As String
//!     Dim result As String
//!     
//!     result = ""
//!     For i = 1 To Len(text)
//!         char = Mid$(text, i, 1)
//!         If (char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or _
//!            (char >= "0" And char <= "9") Then
//!             result = result & char
//!         End If
//!     Next i
//!     
//!     RemovePunctuation = result
//! End Function
//! ```
//!
//! ### Example 2: String Reversal Utilities
//! ```vb6
//! ' Module: StringReversalUtils
//! ' Utilities for reversing strings in various ways
//! Option Explicit
//!
//! Public Function ReverseString(text As String) As String
//!     ReverseString = StrReverse(text)
//! End Function
//!
//! Public Function ReverseWords(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     words = Split(sentence, " ")
//!     result = ""
//!     
//!     For i = UBound(words) To LBound(words) Step -1
//!         If Len(result) > 0 Then result = result & " "
//!         result = result & words(i)
//!     Next i
//!     
//!     ReverseWords = result
//! End Function
//!
//! Public Function ReverseEachWord(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     
//!     words = Split(sentence, " ")
//!     For i = LBound(words) To UBound(words)
//!         words(i) = StrReverse(words(i))
//!     Next i
//!     
//!     ReverseEachWord = Join(words, " ")
//! End Function
//!
//! Public Function ReverseLines(text As String) As String
//!     Dim lines() As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     lines = Split(text, vbCrLf)
//!     result = ""
//!     
//!     For i = UBound(lines) To LBound(lines) Step -1
//!         If Len(result) > 0 Then result = result & vbCrLf
//!         result = result & lines(i)
//!     Next i
//!     
//!     ReverseLines = result
//! End Function
//!
//! Public Function ReverseArray(arr() As String) As String()
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim j As Integer
//!     
//!     ReDim result(LBound(arr) To UBound(arr))
//!     j = LBound(arr)
//!     
//!     For i = UBound(arr) To LBound(arr) Step -1
//!         result(j) = arr(i)
//!         j = j + 1
//!     Next i
//!     
//!     ReverseArray = result
//! End Function
//! ```
//!
//! ### Example 3: Text Transformer Class
//! ```vb6
//! ' Class: TextTransformer
//! ' Performs various text transformations including reversal
//! Option Explicit
//!
//! Public Function Transform(text As String, transformType As String) As String
//!     Select Case UCase$(transformType)
//!         Case "REVERSE"
//!             Transform = StrReverse(text)
//!         Case "REVERSE_WORDS"
//!             Transform = ReverseWords(text)
//!         Case "REVERSE_EACH_WORD"
//!             Transform = ReverseEachWord(text)
//!         Case "MIRROR"
//!             Transform = text & " " & StrReverse(text)
//!         Case "PALINDROME"
//!             Transform = text & StrReverse(text)
//!         Case Else
//!             Transform = text
//!     End Select
//! End Function
//!
//! Private Function ReverseWords(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     words = Split(sentence, " ")
//!     result = ""
//!     
//!     For i = UBound(words) To LBound(words) Step -1
//!         If Len(result) > 0 Then result = result & " "
//!         result = result & words(i)
//!     Next i
//!     
//!     ReverseWords = result
//! End Function
//!
//! Private Function ReverseEachWord(sentence As String) As String
//!     Dim words() As String
//!     Dim i As Integer
//!     
//!     words = Split(sentence, " ")
//!     For i = LBound(words) To UBound(words)
//!         words(i) = StrReverse(words(i))
//!     Next i
//!     
//!     ReverseEachWord = Join(words, " ")
//! End Function
//!
//! Public Function CreatePalindrome(text As String) As String
//!     CreatePalindrome = text & StrReverse(text)
//! End Function
//!
//! Public Function IsPalindrome(text As String) As Boolean
//!     IsPalindrome = (text = StrReverse(text))
//! End Function
//! ```
//!
//! ### Example 4: String Analyzer Module
//! ```vb6
//! ' Module: StringAnalyzer
//! ' Analyzes strings using reversal techniques
//! Option Explicit
//!
//! Public Function ContainsPalindrome(text As String, minLength As Integer) As Boolean
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim substring As String
//!     
//!     For i = 1 To Len(text) - minLength + 1
//!         For j = minLength To Len(text) - i + 1
//!             substring = Mid$(text, i, j)
//!             If substring = StrReverse(substring) Then
//!                 ContainsPalindrome = True
//!                 Exit Function
//!             End If
//!         Next j
//!     Next i
//!     
//!     ContainsPalindrome = False
//! End Function
//!
//! Public Function FindAllPalindromes(text As String, minLength As Integer) As Collection
//!     Dim palindromes As New Collection
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim substring As String
//!     
//!     For i = 1 To Len(text)
//!         For j = minLength To Len(text) - i + 1
//!             substring = Mid$(text, i, j)
//!             If substring = StrReverse(substring) Then
//!                 On Error Resume Next
//!                 palindromes.Add substring, substring
//!                 On Error GoTo 0
//!             End If
//!         Next j
//!     Next i
//!     
//!     Set FindAllPalindromes = palindromes
//! End Function
//!
//! Public Function GetSymmetryScore(text As String) As Double
//!     ' Calculate how symmetric a string is (0-100%)
//!     Dim matches As Integer
//!     Dim i As Integer
//!     Dim len As Integer
//!     
//!     len = Len(text)
//!     If len = 0 Then
//!         GetSymmetryScore = 0
//!         Exit Function
//!     End If
//!     
//!     matches = 0
//!     For i = 1 To len \ 2
//!         If Mid$(text, i, 1) = Mid$(text, len - i + 1, 1) Then
//!             matches = matches + 1
//!         End If
//!     Next i
//!     
//!     GetSymmetryScore = (matches / (len \ 2)) * 100
//! End Function
//!
//! Public Function IsAnagram(str1 As String, str2 As String) As Boolean
//!     ' Not directly using StrReverse, but useful utility
//!     Dim sorted1 As String
//!     Dim sorted2 As String
//!     
//!     sorted1 = SortString(LCase$(str1))
//!     sorted2 = SortString(LCase$(str2))
//!     
//!     IsAnagram = (sorted1 = sorted2)
//! End Function
//!
//! Private Function SortString(text As String) As String
//!     ' Simple bubble sort for demonstration
//!     Dim chars() As String
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim temp As String
//!     
//!     ReDim chars(1 To Len(text))
//!     For i = 1 To Len(text)
//!         chars(i) = Mid$(text, i, 1)
//!     Next i
//!     
//!     For i = 1 To UBound(chars) - 1
//!         For j = i + 1 To UBound(chars)
//!             If chars(i) > chars(j) Then
//!                 temp = chars(i)
//!                 chars(i) = chars(j)
//!                 chars(j) = temp
//!             End If
//!         Next j
//!     Next i
//!     
//!     SortString = Join(chars, "")
//! End Function
//! ```
//!
//! ## Error Handling
//! The `StrReverse` function typically does not raise errors under normal circumstances:
//!
//! - Returns empty string if input is empty string
//! - Returns `Null` if input is `Null` (not an error)
//! - **Error 13 (Type mismatch)**: If `expression` cannot be converted to a string
//!
//! ## Performance Notes
//! - Very fast and efficient (optimized native function)
//! - Much faster than manual character-by-character reversal in VB6
//! - Performance is O(n) where n is string length
//! - No significant overhead for typical string lengths
//! - For very large strings (megabytes), consider memory constraints
//!
//! ## Best Practices
//! 1. **Use `StrReverse`** instead of manual loops for reversing strings (faster and cleaner)
//! 2. **Handle Null values** explicitly when working with Variant types
//! 3. **Normalize before palindrome checks** (remove spaces, convert case)
//! 4. **Don't use for security** - `StrReverse` is not encryption, only obfuscation
//! 5. **Cache reversed strings** if used multiple times in comparisons
//! 6. **Combine with other functions** like `LCase$`, `Trim$` for text processing
//! 7. **Test edge cases** like empty strings and single-character strings
//! 8. **Document intent** when using `StrReverse` in non-obvious ways
//! 9. **Consider alternatives** for word-level reversal (Split/Join approach)
//! 10. **Use for validation** like palindrome checking or symmetry analysis
//!
//! ## Comparison Table
//!
//! | Approach | Code | Speed | Clarity |
//! |----------|------|-------|---------|
//! | `StrReverse` | `StrReverse(s)` | Fast | Excellent |
//! | Manual loop | `For i = Len(s) To 1 Step -1...` | Slow | Poor |
//! | Recursion | `ReverseRecursive(s)` | Very slow | Poor |
//!
//! ## Platform Notes
//! - Available in VB6 and VBA
//! - **Not available in `VBScript`** (must implement manually)
//! - Works correctly with Unicode characters
//! - Behavior consistent across VB6 and VBA
//! - No locale-specific behavior
//!
//! ## Limitations
//! - Cannot reverse only part of a string (use `Mid$` to extract first)
//! - Cannot reverse words rather than characters (use Split/Join)
//! - Not available in `VBScript`
//! - Returns `Null` for `Null` input (may be unexpected)
//! - No option to reverse specific character ranges
//! - Cannot specify custom reversal rules

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_strreverse_basic() {
        let source = r#"
Sub Test()
    result = StrReverse("Hello")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
        assert!(debug.contains("Hello"));
    }

    #[test]
    fn test_strreverse_variable_assignment() {
        let source = r#"
Sub Test()
    Dim reversed As String
    reversed = StrReverse(text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
        assert!(debug.contains("text"));
    }

    #[test]
    fn test_strreverse_palindrome_check() {
        let source = r#"
Sub Test()
    If text = StrReverse(text) Then
        MsgBox "Palindrome"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_function_return() {
        let source = r#"
Function Reverse(s As String) As String
    Reverse = StrReverse(s)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_concatenation() {
        let source = r#"
Sub Test()
    result = text & " | " & StrReverse(text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_with_lcase() {
        let source = r#"
Sub Test()
    If LCase$(text) = StrReverse(LCase$(text)) Then
        isPalindrome = True
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_for_loop() {
        let source = r#"
Sub Test()
    For i = LBound(arr) To UBound(arr)
        arr(i) = StrReverse(arr(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_if_statement() {
        let source = r#"
Sub Test()
    If StrReverse(str1) = str2 Then
        MsgBox "Strings are reverses"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print StrReverse(message)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_msgbox() {
        let source = r#"
Sub Test()
    MsgBox StrReverse("Hello World")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_array_assignment() {
        let source = r#"
Sub Test()
    reversed(i) = StrReverse(original(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessString(StrReverse(input))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_select_case() {
        let source = r#"
Sub Test()
    Select Case StrReverse(code)
        Case "123"
            MsgBox "Match"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_do_while() {
        let source = r#"
Sub Test()
    Do While StrReverse(current) <> target
        current = GetNext()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_do_until() {
        let source = r#"
Sub Test()
    Do Until StrReverse(text) = original
        text = Modify(text)
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_while_wend() {
        let source = r#"
Sub Test()
    While Len(StrReverse(str)) > 10
        str = Trim$(str)
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_iif() {
        let source = r#"
Sub Test()
    result = IIf(reverse, StrReverse(text), text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_with_statement() {
        let source = r#"
Sub Test()
    With obj
        .Text = StrReverse(.Text)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_parentheses() {
        let source = r#"
Sub Test()
    result = (StrReverse(str1) & StrReverse(str2))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    reversed = StrReverse(varText)
    If Err.Number <> 0 Then
        reversed = ""
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_property_assignment() {
        let source = r#"
Sub Test()
    obj.ReversedName = StrReverse(obj.Name)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_comparison() {
        let source = r#"
Sub Test()
    isReverse = (str1 = StrReverse(str2))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_nested() {
        let source = r#"
Sub Test()
    result = StrReverse(StrReverse(text))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_with_trim() {
        let source = r#"
Sub Test()
    result = StrReverse(Trim$(input))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_print_statement() {
        let source = r#"
Sub Test()
    Print #1, StrReverse(data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_class_usage() {
        let source = r#"
Sub Test()
    Set processor = New StringProcessor
    processor.SetText StrReverse(original)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }

    #[test]
    fn test_strreverse_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        result = text
    ElseIf x = 2 Then
        result = StrReverse(text)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("StrReverse"));
    }
}
