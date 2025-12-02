//! # `ChrW` Function
//!
//! Returns a `String` containing the Unicode character associated with the specified character code.
//! The "W" suffix indicates this is the wide (Unicode) version of the `Chr` function.
//!
//! ## Syntax
//!
//! ```vb
//! ChrW(charcode)
//! ```
//!
//! ## Parameters
//!
//! - **charcode**: Required. `Long`. A numeric expression that identifies a Unicode character.
//!   Valid values are -32768 to 65535. However, values -32768 to -1 are treated as 65536 + value.
//!
//! ## Returns
//!
//! Returns a `String` containing a single Unicode character corresponding to the specified code.
//!
//! ## Remarks
//!
//! - `ChrW` is used to return Unicode characters (wide characters).
//! - The W suffix stands for "Wide", distinguishing it from the ANSI `ChrB` function.
//! - Valid Unicode code points range from 0 to 65535 (0x0000 to 0xFFFF) in the Basic Multilingual Plane.
//! - Negative values from -32768 to -1 are converted by adding 65536, allowing access to the full range.
//! - `ChrW` is essential for working with international characters, symbols, and emoji within the BMP.
//! - For ANSI characters (0-255), `ChrW` and `Chr` produce the same results on systems using single-byte character sets.
//! - Characters outside the Basic Multilingual Plane (above 65535) require surrogate pairs in VB6.
//! - If charcode is outside the valid range, a runtime error occurs (Error 5: Invalid procedure call or argument).
//!
//! ## Typical Uses
//!
//! 1. **International text support** - Create strings with characters from various languages
//! 2. **Unicode symbol insertion** - Insert mathematical symbols, currency symbols, arrows, etc.
//! 3. **XML/HTML entity handling** - Convert numeric character references to actual characters
//! 4. **Special character creation** - Generate Unicode control characters and formatting marks
//! 5. **Cross-platform text** - Ensure consistent character representation across different systems
//! 6. **Unicode file processing** - Read and write Unicode text files correctly
//! 7. **Internationalization (i18n)** - Support multiple languages in applications
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple Unicode character
//! Dim ch As String
//! ch = ChrW(65)  ' Returns "A" (same as Chr for ASCII range)
//! ```
//!
//! ```vb
//! ' Example 2: Euro symbol
//! Dim euro As String
//! euro = ChrW(8364)  ' Returns "€"
//! ```
//!
//! ```vb
//! ' Example 3: Greek letter
//! Dim alpha As String
//! alpha = ChrW(945)  ' Returns "α" (Greek small letter alpha)
//! ```
//!
//! ```vb
//! ' Example 4: Chinese character
//! Dim hanzi As String
//! hanzi = ChrW(20013)  ' Returns "中" (Chinese character for "middle")
//! ```
//!
//! ## Common Patterns
//!
//! ### Unicode Line Separator
//! ```vb
//! Const UNICODE_LINE_SEP = 8232
//! text = "Line 1" & ChrW(UNICODE_LINE_SEP) & "Line 2"
//! ```
//!
//! ### Bullet Point List
//! ```vb
//! Dim bullet As String
//! bullet = ChrW(8226)  ' "•"
//! list = bullet & " Item 1" & vbCrLf & bullet & " Item 2"
//! ```
//!
//! ### Copyright Symbol
//! ```vb
//! Dim copyright As String
//! copyright = ChrW(169)  ' "©"
//! notice = "Copyright " & copyright & " 2025"
//! ```
//!
//! ### Mathematical Symbols
//! ```vb
//! Dim infinity As String, pi As String
//! infinity = ChrW(8734)  ' "∞"
//! pi = ChrW(960)  ' "π"
//! equation = pi & " " & ChrW(8776) & " 3.14"  ' "π ≈ 3.14"
//! ```
//!
//! ### Arrow Symbols
//! ```vb
//! Dim rightArrow As String
//! rightArrow = ChrW(8594)  ' "→"
//! leftArrow = ChrW(8592)   ' "←"
//! ```
//!
//! ### Non-Breaking Space
//! ```vb
//! Dim nbsp As String
//! nbsp = ChrW(160)  ' Non-breaking space
//! text = "Price:" & nbsp & "$100"
//! ```
//!
//! ### Emoji and Symbols (BMP only)
//! ```vb
//! Dim heart As String, star As String
//! heart = ChrW(9829)  ' "♥"
//! star = ChrW(9733)   ' "★"
//! ```
//!
//! ### Zero-Width Characters
//! ```vb
//! Dim zwj As String
//! zwj = ChrW(8205)  ' Zero-width joiner
//! ```
//!
//! ### Currency Symbols
//! ```vb
//! Dim pound As String, yen As String
//! pound = ChrW(163)  ' "£"
//! yen = ChrW(165)    ' "¥"
//! ```
//!
//! ### Diacritical Marks
//! ```vb
//! Dim acute As String
//! acute = ChrW(180)  ' "´" (acute accent)
//! combined = "e" & ChrW(769)  ' Combining acute accent
//! ```
//!
//! ## Advanced Examples
//!
//! ### Building Multilingual Text
//! ```vb
//! Function GetGreeting(language As String) As String
//!     Select Case language
//!         Case "chinese"
//!             GetGreeting = ChrW(20320) & ChrW(22909)  ' "你好"
//!         Case "japanese"
//!             GetGreeting = ChrW(12371) & ChrW(12435) & ChrW(12395) & ChrW(12385) & ChrW(12399)  ' "こんにちは"
//!         Case "korean"
//!             GetGreeting = ChrW(50504) & ChrW(45397) & ChrW(54616) & ChrW(49464) & ChrW(50836)  ' "안녕하세요"
//!         Case "russian"
//!             GetGreeting = ChrW(1055) & ChrW(1088) & ChrW(1080) & ChrW(1074) & ChrW(1077) & ChrW(1090)  ' "Привет"
//!         Case Else
//!             GetGreeting = "Hello"
//!     End Select
//! End Function
//! ```
//!
//! ### HTML Entity Decoder
//! ```vb
//! Function DecodeNumericEntity(entity As String) As String
//!     ' Decode &#nnnn; or &#xhhhh; entities
//!     Dim code As Long
//!     
//!     If InStr(entity, "&#x") > 0 Then
//!         ' Hexadecimal entity
//!         code = CLng("&H" & Mid(entity, 4, Len(entity) - 4))
//!     ElseIf InStr(entity, "&#") > 0 Then
//!         ' Decimal entity
//!         code = CLng(Mid(entity, 3, Len(entity) - 3))
//!     End If
//!     
//!     If code >= 0 And code <= 65535 Then
//!         DecodeNumericEntity = ChrW(code)
//!     End If
//! End Function
//! ```
//!
//! ### Unicode Range Validator
//! ```vb
//! Function IsInUnicodeRange(char As String, rangeStart As Long, rangeEnd As Long) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     IsInUnicodeRange = (code >= rangeStart And code <= rangeEnd)
//! End Function
//!
//! ' Example usage:
//! ' If IsInUnicodeRange(char, 0x4E00, 0x9FFF) Then
//! '     Debug.Print "CJK Unified Ideograph"
//! ' End If
//! ```
//!
//! ### Unicode Text File Writer
//! ```vb
//! Sub WriteUnicodeFile(filename As String, text As String)
//!     Dim fso As Object, stream As Object
//!     Set fso = CreateObject("Scripting.FileSystemObject")
//!     Set stream = CreateObject("ADODB.Stream")
//!     
//!     stream.Type = 2  ' Text
//!     stream.Charset = "UTF-8"
//!     stream.Open
//!     
//!     ' Add BOM for UTF-8
//!     stream.WriteText ChrW(65279)  ' UTF-8 BOM
//!     stream.WriteText text
//!     
//!     stream.SaveToFile filename, 2
//!     stream.Close
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChrW(charcode As Long) As String
//!     On Error GoTo ErrorHandler
//!     
//!     ' Normalize negative values
//!     If charcode < 0 Then
//!         charcode = 65536 + charcode
//!     End If
//!     
//!     If charcode >= 0 And charcode <= 65535 Then
//!         SafeChrW = ChrW(charcode)
//!     Else
//!         SafeChrW = "?"  ' Replacement character for invalid codes
//!     End If
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeChrW = "?"
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `ChrW` is a fast operation with minimal overhead
//! - When building long Unicode strings, use string concatenation efficiently or a `StringBuilder` pattern
//! - For repeated character creation, consider caching the result
//! - `AscW` is the inverse function of `ChrW`
//!
//! ## Best Practices
//!
//! 1. **Use `ChrW` for Unicode** - Always use `ChrW` instead of `Chr` when working with international text
//! 2. **Validate ranges** - Check that character codes are within valid Unicode ranges
//! 3. **Handle errors** - Wrap `ChrW` calls in error handlers when processing user input
//! 4. **Document character codes** - Use constants or comments to explain non-obvious character codes
//! 5. **Test with actual data** - Verify Unicode text displays correctly in target environments
//! 6. **Consider normalization** - Be aware that some characters can be represented multiple ways
//! 7. **Use UTF-8 for files** - When saving Unicode text, prefer UTF-8 encoding
//!
//! ## Comparison with Related Functions
//!
//! | Function | Character Set | Range | Use Case |
//! |----------|---------------|-------|----------|
//! | `Chr` | System default (ANSI/Unicode) | 0-255 | Legacy code, simple ASCII |
//! | `ChrB` | ANSI (byte) | 0-255 | ANSI text, byte operations |
//! | `ChrW` | Unicode (wide) | 0-65535 | International text, symbols |
//! | `AscW` | Unicode (inverse) | Returns 0-65535 | Get Unicode code from character |
//!
//! ## Unicode Ranges Reference
//!
//! Some common Unicode ranges that work with `ChrW`:
//!
//! - **Basic Latin**: 0-127 (ASCII)
//! - **Latin-1 Supplement**: 128-255
//! - **Greek and Coptic**: 880-1023
//! - **Cyrillic**: 1024-1279
//! - **Hebrew**: 1424-1535
//! - **Arabic**: 1536-1791
//! - **CJK Unified Ideographs**: 19968-40959
//! - **Hangul Syllables**: 44032-55203
//! - **Currency Symbols**: 8352-8399
//! - **Mathematical Operators**: 8704-8959
//! - **Arrows**: 8592-8703
//! - **Box Drawing**: 9472-9599
//!
//! ## Platform Notes
//!
//! - VB6 supports Unicode through `ChrW` but stores strings internally in the system's native format
//! - On Windows NT-based systems, strings are stored as Unicode (UTF-16)
//! - On Windows 95/98/ME, strings use ANSI encoding, which may cause issues with characters outside the current code page
//! - When running on older systems, test thoroughly with non-ASCII characters
//! - Modern Windows systems (XP and later) handle Unicode properly
//!
//! ## Limitations
//!
//! - `ChrW` only supports the Basic Multilingual Plane (BMP), code points 0-65535
//! - Characters outside the BMP (like some emoji) require surrogate pairs in VB6
//! - Surrogate pairs are complex and require combining two `ChrW` calls
//! - Not all fonts support all Unicode characters; display depends on available fonts
//! - Some Unicode features like combining characters may not render correctly in all VB6 controls

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn chrw_simple_ascii() {
        let source = r#"
Sub Test()
    ch = ChrW(65)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrw_euro_symbol() {
        let source = r#"
Sub Test()
    euro = ChrW(8364)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrw_greek_letter() {
        let source = r#"
Sub Test()
    alpha = ChrW(945)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrw_chinese_character() {
        let source = r#"
Sub Test()
    hanzi = ChrW(20013)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrw_line_separator() {
        let source = r#"
Sub Test()
    text = "Line 1" & ChrW(8232) & "Line 2"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_bullet_point() {
        let source = r#"
Sub Test()
    bullet = ChrW(8226)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_copyright_symbol() {
        let source = r#"
Sub Test()
    copyright = ChrW(169)
    notice = "Copyright " & copyright & " 2025"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_mathematical_symbols() {
        let source = r#"
Sub Test()
    infinity = ChrW(8734)
    pi = ChrW(960)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_arrow_symbols() {
        let source = r#"
Sub Test()
    rightArrow = ChrW(8594)
    leftArrow = ChrW(8592)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_non_breaking_space() {
        let source = r#"
Sub Test()
    nbsp = ChrW(160)
    text = "Price:" & nbsp & "$100"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_heart_symbol() {
        let source = r#"
Sub Test()
    heart = ChrW(9829)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_currency_symbols() {
        let source = r#"
Sub Test()
    pound = ChrW(163)
    yen = ChrW(165)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_multilingual_function() {
        let source = r#"
Function GetGreeting(language As String) As String
    Select Case language
        Case "chinese"
            GetGreeting = ChrW(20320) & ChrW(22909)
        Case Else
            GetGreeting = "Hello"
    End Select
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_with_variable() {
        let source = r#"
Sub Test()
    Dim code As Long
    code = 945
    ch = ChrW(code)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_in_loop() {
        let source = r#"
Sub Test()
    For i = 65 To 90
        chars = chars & ChrW(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_range_validator() {
        let source = r#"
Function IsInUnicodeRange(code As Long, rangeStart As Long, rangeEnd As Long) As Boolean
    IsInUnicodeRange = (code >= rangeStart And code <= rangeEnd)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_safe_wrapper() {
        let source = r#"
Function SafeChrW(charcode As Long) As String
    If charcode >= 0 And charcode <= 65535 Then
        SafeChrW = ChrW(charcode)
    Else
        SafeChrW = "?"
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_bom_marker() {
        let source = r#"
Sub Test()
    bom = ChrW(65279)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_combining_accent() {
        let source = r#"
Sub Test()
    acute = ChrW(180)
    combined = "e" & ChrW(769)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrw_in_conditional() {
        let source = r#"
Sub Test()
    If needsSpecial Then
        text = ChrW(8594)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }
}
