//! VB6 `StrConv` Function
//!
//! The `StrConv` function converts a string to a specified format.
//!
//! ## Syntax
//! ```vb6
//! StrConv(string, conversion[, LCID])
//! ```
//!
//! ## Parameters
//! - `string`: Required. String expression to be converted.
//! - `conversion`: Required. Integer specifying the type of conversion to perform. Can be one or more of the following constants (combined with `+` or `Or`):
//!   - `vbUpperCase` (1): Converts the string to uppercase characters
//!   - `vbLowerCase` (2): Converts the string to lowercase characters
//!   - `vbProperCase` (3): Converts the first letter of every word to uppercase
//!   - `vbWide` (4): Converts narrow (single-byte) characters to wide (double-byte) characters
//!   - `vbNarrow` (8): Converts wide (double-byte) characters to narrow (single-byte) characters
//!   - `vbKatakana` (16): Converts Hiragana characters to Katakana characters (Japanese)
//!   - `vbHiragana` (32): Converts Katakana characters to Hiragana characters (Japanese)
//!   - `vbUnicode` (64): Converts the string to Unicode using the default code page
//!   - `vbFromUnicode` (128): Converts the string from Unicode to the default code page
//! - `LCID`: Optional. `LocaleID` value, if different from the system `LocaleID` value. Default is the system `LocaleID`.
//!
//! ## Returns
//! Returns a `Variant` (String) containing the converted string, or a `Variant` (Byte array) when converting to/from Unicode.
//!
//! ## Remarks
//! The `StrConv` function provides powerful string transformation capabilities:
//!
//! - **Case conversion**: `vbUpperCase`, `vbLowerCase`, and `vbProperCase` for text normalization
//! - **Proper case rules**: `vbProperCase` capitalizes first letter after spaces and certain punctuation
//! - **Wide/Narrow conversion**: For Asian languages with double-byte character sets (DBCS)
//! - **Japanese character conversion**: `vbKatakana` and `vbHiragana` for Japanese text
//! - **Unicode conversion**: `vbUnicode` converts string to byte array, `vbFromUnicode` converts byte array to string
//! - **Combining conversions**: Can combine multiple conversions using `+` or `Or` (e.g., `vbUpperCase + vbWide`)
//! - **Return type varies**: String conversions return String, Unicode conversions return Byte array
//! - **Locale-aware**: Proper case conversion respects locale settings
//! - **Performance**: Efficient for bulk text transformation
//!
//! ### Case Conversion Details
//! - `vbUpperCase`: Converts all characters to uppercase using locale rules
//! - `vbLowerCase`: Converts all characters to lowercase using locale rules
//! - `vbProperCase`: Capitalizes first letter of each word, converts rest to lowercase
//!   - Words are delimited by spaces, tabs, and some punctuation
//!   - Preserves existing spacing and punctuation
//!
//! ### Unicode Conversion Details
//! - `vbUnicode`: Converts String to Byte array containing Unicode (UTF-16LE) representation
//!   - Each character becomes 2 bytes
//!   - Useful for binary file operations or API calls
//! - `vbFromUnicode`: Converts Byte array back to String
//!   - Expects byte array in Unicode format
//!   - Reverses `vbUnicode` conversion
//!
//! ### Wide/Narrow Conversion (DBCS)
//! - Relevant for Asian languages (Japanese, Chinese, Korean)
//! - Wide characters occupy two bytes, narrow characters occupy one byte
//! - `vbWide`: Converts half-width to full-width characters
//! - `vbNarrow`: Converts full-width to half-width characters
//!
//! ## Typical Uses
//! 1. **Text Normalization**: Convert user input to consistent case for comparisons
//! 2. **Title Formatting**: Format text in proper case for titles and headings
//! 3. **Data Validation**: Normalize strings before validation or storage
//! 4. **Case-Insensitive Operations**: Convert to uppercase/lowercase for comparisons
//! 5. **Unicode File I/O**: Convert strings to/from Unicode byte arrays for file operations
//! 6. **API Calls**: Convert strings to Unicode byte arrays for Win32 API calls
//! 7. **Japanese Text Processing**: Convert between Hiragana and Katakana
//! 8. **Database Storage**: Normalize case before storing in databases
//!
//! ## Basic Examples
//!
//! ### Example 1: Case Conversion
//! ```vb6
//! Dim text As String
//! Dim result As String
//!
//! text = "Hello World"
//!
//! result = StrConv(text, vbUpperCase)    ' "HELLO WORLD"
//! result = StrConv(text, vbLowerCase)    ' "hello world"
//! result = StrConv(text, vbProperCase)   ' "Hello World"
//!
//! text = "the quick brown fox"
//! result = StrConv(text, vbProperCase)   ' "The Quick Brown Fox"
//! ```
//!
//! ### Example 2: Unicode Conversion
//! ```vb6
//! Dim text As String
//! Dim bytes() As Byte
//! Dim restored As String
//!
//! text = "Hello"
//!
//! ' Convert to Unicode byte array
//! bytes = StrConv(text, vbUnicode)
//! ' bytes contains: 72, 0, 101, 0, 108, 0, 108, 0, 111, 0
//!
//! ' Convert back to string
//! restored = StrConv(bytes, vbFromUnicode)  ' "Hello"
//! ```
//!
//! ### Example 3: Combining Conversions
//! ```vb6
//! Dim text As String
//! Dim result As String
//!
//! text = "hello"
//!
//! ' Combine uppercase with wide conversion (for DBCS)
//! result = StrConv(text, vbUpperCase + vbWide)
//! ```
//!
//! ### Example 4: Proper Case Formatting
//! ```vb6
//! Dim name As String
//! Dim formatted As String
//!
//! name = "JOHN Q. PUBLIC"
//! formatted = StrConv(name, vbProperCase)  ' "John Q. Public"
//!
//! name = "o'brien"
//! formatted = StrConv(name, vbProperCase)  ' "O'brien"
//! ' Note: StrConv doesn't handle special cases like O'Brien
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Normalize User Input
//! ```vb6
//! Function NormalizeInput(userInput As String) As String
//!     ' Convert to uppercase for case-insensitive processing
//!     NormalizeInput = StrConv(Trim$(userInput), vbUpperCase)
//! End Function
//! ```
//!
//! ### Pattern 2: Format Name Properly
//! ```vb6
//! Function FormatName(name As String) As String
//!     ' Convert to proper case for display
//!     FormatName = StrConv(Trim$(name), vbProperCase)
//! End Function
//! ```
//!
//! ### Pattern 3: Case-Insensitive Comparison
//! ```vb6
//! Function EqualsIgnoreCase(str1 As String, str2 As String) As Boolean
//!     EqualsIgnoreCase = (StrConv(str1, vbUpperCase) = StrConv(str2, vbUpperCase))
//! End Function
//! ```
//!
//! ### Pattern 4: Write Unicode File
//! ```vb6
//! Sub WriteUnicodeFile(filename As String, text As String)
//!     Dim fileNum As Integer
//!     Dim bytes() As Byte
//!     
//!     bytes = StrConv(text, vbUnicode)
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     Put #fileNum, , bytes
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Pattern 5: Read Unicode File
//! ```vb6
//! Function ReadUnicodeFile(filename As String) As String
//!     Dim fileNum As Integer
//!     Dim bytes() As Byte
//!     Dim fileSize As Long
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     
//!     fileSize = LOF(fileNum)
//!     ReDim bytes(0 To fileSize - 1)
//!     Get #fileNum, , bytes
//!     Close #fileNum
//!     
//!     ReadUnicodeFile = StrConv(bytes, vbFromUnicode)
//! End Function
//! ```
//!
//! ### Pattern 6: Convert Array of Strings
//! ```vb6
//! Sub ConvertArrayToUpperCase(arr() As String)
//!     Dim i As Integer
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = StrConv(arr(i), vbUpperCase)
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 7: Title Case for Sentences
//! ```vb6
//! Function FormatTitle(title As String) As String
//!     Dim result As String
//!     
//!     ' Convert to proper case
//!     result = StrConv(title, vbProperCase)
//!     
//!     ' Handle articles and prepositions (simplified)
//!     result = Replace(result, " A ", " a ")
//!     result = Replace(result, " An ", " an ")
//!     result = Replace(result, " The ", " the ")
//!     result = Replace(result, " Of ", " of ")
//!     result = Replace(result, " In ", " in ")
//!     
//!     FormatTitle = result
//! End Function
//! ```
//!
//! ### Pattern 8: Database Normalization
//! ```vb6
//! Function NormalizeForDatabase(value As String) As String
//!     ' Trim and convert to uppercase for storage
//!     NormalizeForDatabase = StrConv(Trim$(value), vbUpperCase)
//! End Function
//! ```
//!
//! ### Pattern 9: Compare with Wildcard
//! ```vb6
//! Function MatchesPattern(text As String, pattern As String) As Boolean
//!     ' Case-insensitive pattern matching
//!     MatchesPattern = (StrConv(text, vbUpperCase) Like StrConv(pattern, vbUpperCase))
//! End Function
//! ```
//!
//! ### Pattern 10: Extract Unicode Bytes
//! ```vb6
//! Function GetUnicodeBytes(text As String) As String
//!     Dim bytes() As Byte
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     bytes = StrConv(text, vbUnicode)
//!     
//!     result = ""
//!     For i = LBound(bytes) To UBound(bytes)
//!         result = result & CStr(bytes(i)) & " "
//!     Next i
//!     
//!     GetUnicodeBytes = Trim$(result)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Text Normalizer Class
//! ```vb6
//! ' Class: TextNormalizer
//! ' Provides text normalization and conversion utilities
//! Option Explicit
//!
//! Public Enum NormalizationMode
//!     UpperCase = 1
//!     LowerCase = 2
//!     ProperCase = 3
//!     NoChange = 0
//! End Enum
//!
//! Private m_Mode As NormalizationMode
//! Private m_TrimSpaces As Boolean
//!
//! Public Sub Initialize(mode As NormalizationMode, trimSpaces As Boolean)
//!     m_Mode = mode
//!     m_TrimSpaces = trimSpaces
//! End Sub
//!
//! Public Function Normalize(text As String) As String
//!     Dim result As String
//!     
//!     result = text
//!     
//!     ' Trim if requested
//!     If m_TrimSpaces Then
//!         result = Trim$(result)
//!     End If
//!     
//!     ' Apply case conversion
//!     Select Case m_Mode
//!         Case UpperCase
//!             result = StrConv(result, vbUpperCase)
//!         Case LowerCase
//!             result = StrConv(result, vbLowerCase)
//!         Case ProperCase
//!             result = StrConv(result, vbProperCase)
//!     End Select
//!     
//!     Normalize = result
//! End Function
//!
//! Public Function NormalizeArray(arr() As String) As String()
//!     Dim result() As String
//!     Dim i As Integer
//!     
//!     ReDim result(LBound(arr) To UBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         result(i) = Normalize(arr(i))
//!     Next i
//!     
//!     NormalizeArray = result
//! End Function
//!
//! Public Function NormalizeCollection(col As Collection) As Collection
//!     Dim result As New Collection
//!     Dim item As Variant
//!     
//!     For Each item In col
//!         result.Add Normalize(CStr(item))
//!     Next item
//!     
//!     Set NormalizeCollection = result
//! End Function
//! ```
//!
//! ### Example 2: Unicode File Handler
//! ```vb6
//! ' Class: UnicodeFileHandler
//! ' Handles reading and writing Unicode text files
//! Option Explicit
//!
//! Public Sub WriteFile(filename As String, text As String, Optional appendMode As Boolean = False)
//!     Dim fileNum As Integer
//!     Dim bytes() As Byte
//!     Dim existingBytes() As Byte
//!     Dim existingSize As Long
//!     Dim newSize As Long
//!     
//!     bytes = StrConv(text, vbUnicode)
//!     fileNum = FreeFile
//!     
//!     If appendMode And Dir(filename) <> "" Then
//!         ' Read existing content
//!         Open filename For Binary As #fileNum
//!         existingSize = LOF(fileNum)
//!         If existingSize > 0 Then
//!             ReDim existingBytes(0 To existingSize - 1)
//!             Get #fileNum, , existingBytes
//!         End If
//!         Close #fileNum
//!         
//!         ' Combine existing and new bytes
//!         newSize = existingSize + UBound(bytes) + 1
//!         ReDim Preserve existingBytes(0 To newSize - 1)
//!         
//!         Dim i As Long
//!         For i = 0 To UBound(bytes)
//!             existingBytes(existingSize + i) = bytes(i)
//!         Next i
//!         
//!         bytes = existingBytes
//!     End If
//!     
//!     ' Write to file
//!     Open filename For Binary As #fileNum
//!     Put #fileNum, , bytes
//!     Close #fileNum
//! End Sub
//!
//! Public Function ReadFile(filename As String) As String
//!     Dim fileNum As Integer
//!     Dim bytes() As Byte
//!     Dim fileSize As Long
//!     
//!     If Dir(filename) = "" Then
//!         Err.Raise 53, , "File not found"
//!     End If
//!     
//!     fileNum = FreeFile
//!     Open filename For Binary As #fileNum
//!     
//!     fileSize = LOF(fileNum)
//!     If fileSize = 0 Then
//!         Close #fileNum
//!         ReadFile = ""
//!         Exit Function
//!     End If
//!     
//!     ReDim bytes(0 To fileSize - 1)
//!     Get #fileNum, , bytes
//!     Close #fileNum
//!     
//!     ReadFile = StrConv(bytes, vbFromUnicode)
//! End Function
//!
//! Public Function ReadLines(filename As String) As String()
//!     Dim content As String
//!     Dim lines() As String
//!     
//!     content = ReadFile(filename)
//!     lines = Split(content, vbCrLf)
//!     
//!     ReadLines = lines
//! End Function
//!
//! Public Sub WriteLines(filename As String, lines() As String)
//!     Dim content As String
//!     content = Join(lines, vbCrLf)
//!     WriteFile filename, content
//! End Sub
//! ```
//!
//! ### Example 3: Case Converter Module
//! ```vb6
//! ' Module: CaseConverter
//! ' Utilities for case conversion and formatting
//! Option Explicit
//!
//! Public Function ToUpper(text As String) As String
//!     ToUpper = StrConv(text, vbUpperCase)
//! End Function
//!
//! Public Function ToLower(text As String) As String
//!     ToLower = StrConv(text, vbLowerCase)
//! End Function
//!
//! Public Function ToProper(text As String) As String
//!     ToProper = StrConv(text, vbProperCase)
//! End Function
//!
//! Public Function ToTitleCase(text As String) As String
//!     ' More sophisticated title case
//!     Dim result As String
//!     Dim words() As String
//!     Dim i As Integer
//!     Dim lowercaseWords As String
//!     
//!     ' Start with proper case
//!     result = StrConv(text, vbProperCase)
//!     
//!     ' Lowercase certain words (not first word)
//!     lowercaseWords = " a an the and but or for nor of in on at to from by "
//!     words = Split(result, " ")
//!     
//!     For i = 1 To UBound(words)  ' Start at 1 to skip first word
//!         If InStr(lowercaseWords, " " & LCase$(words(i)) & " ") > 0 Then
//!             words(i) = LCase$(words(i))
//!         End If
//!     Next i
//!     
//!     ToTitleCase = Join(words, " ")
//! End Function
//!
//! Public Function ToggleCase(text As String) As String
//!     ' Toggle case of each character
//!     Dim i As Integer
//!     Dim char As String
//!     Dim result As String
//!     
//!     result = ""
//!     For i = 1 To Len(text)
//!         char = Mid$(text, i, 1)
//!         If char = UCase$(char) Then
//!             result = result & LCase$(char)
//!         Else
//!             result = result & UCase$(char)
//!         End If
//!     Next i
//!     
//!     ToggleCase = result
//! End Function
//!
//! Public Function IsAllUpper(text As String) As Boolean
//!     IsAllUpper = (text = StrConv(text, vbUpperCase))
//! End Function
//!
//! Public Function IsAllLower(text As String) As Boolean
//!     IsAllLower = (text = StrConv(text, vbLowerCase))
//! End Function
//!
//! Public Function IsProperCase(text As String) As Boolean
//!     IsProperCase = (text = StrConv(text, vbProperCase))
//! End Function
//! ```
//!
//! ### Example 4: String Comparison Helper
//! ```vb6
//! ' Module: StringComparisonHelper
//! ' Case-insensitive comparison utilities
//! Option Explicit
//!
//! Public Function EqualsIgnoreCase(str1 As String, str2 As String) As Boolean
//!     EqualsIgnoreCase = (StrConv(str1, vbUpperCase) = StrConv(str2, vbUpperCase))
//! End Function
//!
//! Public Function StartsWithIgnoreCase(text As String, prefix As String) As Boolean
//!     Dim textUpper As String
//!     Dim prefixUpper As String
//!     
//!     textUpper = StrConv(text, vbUpperCase)
//!     prefixUpper = StrConv(prefix, vbUpperCase)
//!     
//!     StartsWithIgnoreCase = (Left$(textUpper, Len(prefixUpper)) = prefixUpper)
//! End Function
//!
//! Public Function EndsWithIgnoreCase(text As String, suffix As String) As Boolean
//!     Dim textUpper As String
//!     Dim suffixUpper As String
//!     
//!     textUpper = StrConv(text, vbUpperCase)
//!     suffixUpper = StrConv(suffix, vbUpperCase)
//!     
//!     EndsWithIgnoreCase = (Right$(textUpper, Len(suffixUpper)) = suffixUpper)
//! End Function
//!
//! Public Function ContainsIgnoreCase(text As String, searchValue As String) As Boolean
//!     Dim textUpper As String
//!     Dim searchUpper As String
//!     
//!     textUpper = StrConv(text, vbUpperCase)
//!     searchUpper = StrConv(searchValue, vbUpperCase)
//!     
//!     ContainsIgnoreCase = (InStr(textUpper, searchUpper) > 0)
//! End Function
//!
//! Public Function IndexOfIgnoreCase(text As String, searchValue As String) As Long
//!     Dim textUpper As String
//!     Dim searchUpper As String
//!     
//!     textUpper = StrConv(text, vbUpperCase)
//!     searchUpper = StrConv(searchValue, vbUpperCase)
//!     
//!     IndexOfIgnoreCase = InStr(textUpper, searchUpper)
//! End Function
//!
//! Public Function ReplaceIgnoreCase(text As String, findText As String, _
//!                                   replaceText As String) As String
//!     Dim result As String
//!     Dim pos As Long
//!     Dim lastPos As Long
//!     Dim textUpper As String
//!     Dim findUpper As String
//!     
//!     result = ""
//!     lastPos = 1
//!     textUpper = StrConv(text, vbUpperCase)
//!     findUpper = StrConv(findText, vbUpperCase)
//!     
//!     pos = InStr(lastPos, textUpper, findUpper)
//!     Do While pos > 0
//!         result = result & Mid$(text, lastPos, pos - lastPos) & replaceText
//!         lastPos = pos + Len(findText)
//!         pos = InStr(lastPos, textUpper, findUpper)
//!     Loop
//!     
//!     result = result & Mid$(text, lastPos)
//!     ReplaceIgnoreCase = result
//! End Function
//! ```
//!
//! ## Error Handling
//! The `StrConv` function can raise the following errors:
//!
//! - **Error 5 (Invalid procedure call or argument)**: If `conversion` constant is invalid or incompatible combinations are used
//! - **Error 13 (Type mismatch)**: If `string` argument cannot be converted to a string
//! - **Error 6 (Overflow)**: In rare cases with very large strings or byte arrays
//!
//! ## Performance Notes
//! - Very fast for case conversions (uppercase, lowercase, proper case)
//! - Unicode conversions are efficient but create byte arrays (memory overhead)
//! - Proper case conversion slower than upper/lower case (more complex rules)
//! - Wide/Narrow conversions only relevant for DBCS environments
//! - Consider caching converted values if used repeatedly
//! - For simple uppercase/lowercase, `UCase$` and `LCase$` may be slightly faster
//!
//! ## Best Practices
//! 1. **Use for normalization** before comparisons or storage
//! 2. **Combine with Trim$** to remove leading/trailing spaces before conversion
//! 3. **Handle proper case limitations** - doesn't handle special cases like "O'Brien" or "`McDonald`"
//! 4. **Cache conversion constants** in variables for clarity (e.g., `Const UPPER_CASE = vbUpperCase`)
//! 5. **Use Unicode conversion** for binary file I/O or API calls requiring Unicode
//! 6. **Test with locale-specific text** when using proper case or locale-dependent conversions
//! 7. **Document conversion type** in comments when not obvious from context
//! 8. **Consider alternatives** - `UCase$`, `LCase$` for simple case conversion (slightly faster)
//! 9. **Validate conversion parameter** when accepting user input
//! 10. **Handle byte array return** appropriately when using `vbUnicode` conversion
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Returns | Locale-Aware |
//! |----------|---------|---------|--------------|
//! | `StrConv` | Multiple conversions | String or Byte array | Yes |
//! | `UCase$` | Uppercase only | String | Yes |
//! | `LCase$` | Lowercase only | String | Yes |
//! | `Format$` | General formatting | String | Yes |
//!
//! ## Platform Notes
//! - Available in VB6 and VBA
//! - Not available in `VBScript`
//! - `vbWide`/`vbNarrow` primarily for Asian language environments
//! - `vbKatakana`/`vbHiragana` only meaningful for Japanese text
//! - Unicode conversion uses UTF-16LE (Windows default)
//! - Proper case rules may vary by locale
//! - LCID parameter rarely used (defaults to system locale)
//!
//! ## Limitations
//! - Proper case doesn't handle special cases (O'Brien, `McDonald`, etc.)
//! - Cannot specify custom word delimiters for proper case
//! - Unicode conversion always uses UTF-16LE (cannot specify encoding)
//! - No direct support for other Unicode formats (UTF-8, UTF-32)
//! - Wide/Narrow conversion limited to DBCS environments
//! - Cannot combine incompatible conversions (e.g., `vbUpperCase + vbLowerCase`)
//! - No validation of byte array format when using `vbFromUnicode`
//! - LCID parameter has limited practical use

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn strconv_basic() {
        let source = r#"
Sub Test()
    result = StrConv("Hello", vbUpperCase)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbUpperCase"),
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
    fn strconv_variable_assignment() {
        let source = r"
Sub Test()
    Dim result As String
    result = StrConv(text, vbUpperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbUpperCase"),
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
    fn strconv_lowercase() {
        let source = r"
Sub Test()
    result = StrConv(input, vbLowerCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        InputKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbLowerCase"),
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
    fn strconv_propercase() {
        let source = r"
Sub Test()
    result = StrConv(name, vbProperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        NameKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbProperCase"),
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
    fn strconv_unicode() {
        let source = r"
Sub Test()
    Dim bytes() As Byte
    bytes = StrConv(text, vbUnicode)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("bytes"),
                        LeftParenthesis,
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        ByteKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bytes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbUnicode"),
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
    fn strconv_from_unicode() {
        let source = r"
Sub Test()
    result = StrConv(byteArray, vbFromUnicode)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("byteArray"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbFromUnicode"),
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
    fn strconv_if_statement() {
        let source = r#"
Sub Test()
    If StrConv(input, vbUpperCase) = "YES" Then
        MsgBox "Confirmed"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            InputKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"YES\""),
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
                                StringLiteral ("\"Confirmed\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn strconv_for_loop() {
        let source = r"
Sub Test()
    For i = LBound(arr) To UBound(arr)
        arr(i) = StrConv(arr(i), vbUpperCase)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LBound"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("arr"),
                                    },
                                },
                            },
                            RightParenthesis,
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
                                        Identifier ("arr"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                CallExpression {
                                    Identifier ("arr"),
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
                                    Identifier ("StrConv"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("arr"),
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
                                            IdentifierExpression {
                                                Identifier ("vbUpperCase"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
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
    fn strconv_function_return() {
        let source = r"
Function ToUpper(text As String) As String
    ToUpper = StrConv(text, vbUpperCase)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ToUpper"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ToUpper"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbUpperCase"),
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
    fn strconv_comparison() {
        let source = r#"
Sub Test()
    If StrConv(str1, vbUpperCase) = StrConv(str2, vbUpperCase) Then
        MsgBox "Equal"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("str1"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("str2"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
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
                                StringLiteral ("\"Equal\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn strconv_with_trim() {
        let source = r"
Sub Test()
    result = StrConv(Trim$(input), vbProperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    InputKeyword,
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
                                        Identifier ("vbProperCase"),
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
    fn strconv_select_case() {
        let source = r#"
Sub Test()
    Select Case StrConv(command, vbUpperCase)
        Case "QUIT"
            Exit Sub
        Case "HELP"
            ShowHelp
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("command"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbUpperCase"),
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
                            StringLiteral ("\"QUIT\""),
                            Newline,
                            StatementList {
                                ExitStatement {
                                    Whitespace,
                                    ExitKeyword,
                                    Whitespace,
                                    SubKeyword,
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"HELP\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("ShowHelp"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
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
    fn strconv_array_assignment() {
        let source = r"
Sub Test()
    normalized(i) = StrConv(original(i), vbUpperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("normalized"),
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("original"),
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
                                    IdentifierExpression {
                                        Identifier ("vbUpperCase"),
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
    fn strconv_function_argument() {
        let source = r"
Sub Test()
    Call ProcessText(StrConv(input, vbProperCase))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    CallStatement {
                        Whitespace,
                        CallKeyword,
                        Whitespace,
                        Identifier ("ProcessText"),
                        LeftParenthesis,
                        Identifier ("StrConv"),
                        LeftParenthesis,
                        InputKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("vbProperCase"),
                        RightParenthesis,
                        RightParenthesis,
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
    fn strconv_concatenation() {
        let source = r#"
Sub Test()
    message = "Hello " & StrConv(name, vbProperCase)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("message"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Hello \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            NameKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbProperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
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
    fn strconv_do_while() {
        let source = r#"
Sub Test()
    Do While StrConv(input, vbUpperCase) <> "DONE"
        input = GetInput()
    Loop
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            InputKeyword,
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"DONE\""),
                            },
                        },
                        Newline,
                        StatementList {
                            InputStatement {
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                Identifier ("GetInput"),
                                LeftParenthesis,
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn strconv_do_until() {
        let source = r#"
Sub Test()
    Do Until StrConv(status, vbUpperCase) = "READY"
        Wait 100
    Loop
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("status"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"READY\""),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Wait"),
                                Whitespace,
                                IntegerLiteral ("100"),
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn strconv_while_wend() {
        let source = r#"
Sub Test()
    While StrConv(cmd, vbUpperCase) <> "EXIT"
        cmd = ProcessCommand()
    Wend
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("StrConv"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("cmd"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"EXIT\""),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("cmd"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("ProcessCommand"),
                                    LeftParenthesis,
                                    ArgumentList,
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn strconv_iif() {
        let source = r"
Sub Test()
    result = IIf(mode = 1, StrConv(text, vbUpperCase), StrConv(text, vbLowerCase))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("mode"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("StrConv"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    TextKeyword,
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("vbUpperCase"),
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
                                        Identifier ("StrConv"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    TextKeyword,
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("vbLowerCase"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn strconv_with_statement() {
        let source = r"
Sub Test()
    With obj
        .Name = StrConv(.Name, vbProperCase)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        NameKeyword,
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("StrConv"),
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
                            NameStatement {
                                NameKeyword,
                                Comma,
                                Whitespace,
                                Identifier ("vbProperCase"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
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
    fn strconv_parentheses() {
        let source = r"
Sub Test()
    result = (StrConv(str1, vbUpperCase) = StrConv(str2, vbUpperCase))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("StrConv"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("str1"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("vbUpperCase"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("StrConv"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("str2"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("vbUpperCase"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn strconv_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = StrConv(varValue, vbProperCase)
    If Err.Number <> 0 Then
        result = ""
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
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
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("varValue"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbProperCase"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("result"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\"\""),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn strconv_property_assignment() {
        let source = r"
Sub Test()
    obj.Title = StrConv(rawTitle, vbProperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("obj"),
                            PeriodOperator,
                            Identifier ("Title"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawTitle"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbProperCase"),
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
    fn strconv_msgbox() {
        let source = r#"
Sub Test()
    MsgBox StrConv("warning: system error", vbProperCase)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        Identifier ("StrConv"),
                        LeftParenthesis,
                        StringLiteral ("\"warning: system error\""),
                        Comma,
                        Whitespace,
                        Identifier ("vbProperCase"),
                        RightParenthesis,
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
    fn strconv_debug_print() {
        let source = r"
Sub Test()
    Debug.Print StrConv(output, vbUpperCase)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("StrConv"),
                        LeftParenthesis,
                        OutputKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("vbUpperCase"),
                        RightParenthesis,
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
    fn strconv_numeric_constant() {
        let source = r"
Sub Test()
    result = StrConv(text, 1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
    fn strconv_combined_conversion() {
        let source = r"
Sub Test()
    result = StrConv(text, vbUpperCase + vbWide)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
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
                            Identifier ("StrConv"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("vbUpperCase"),
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("vbWide"),
                                        },
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
}
