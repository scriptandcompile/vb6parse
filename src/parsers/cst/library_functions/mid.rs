//! # Mid Function
//!
//! Returns a Variant (String) containing a specified number of characters from a string.
//!
//! ## Syntax
//!
//! ```vb
//! Mid(string, start, [length])
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which characters are returned
//!   - Can be any valid string expression
//!   - If Null, returns Null
//!   - If empty string, returns empty string
//!
//! - `start` (Required): Long. Character position in string where the desired part begins
//!   - Uses 1-based indexing (first character is position 1)
//!   - If start > Len(string), returns empty string
//!   - If start < 1, error 5 (Invalid procedure call or argument)
//!
//! - `length` (Optional): Long. Number of characters to return
//!   - If omitted or > characters available, returns all characters from start to end
//!   - If length < 0, error 5 (Invalid procedure call or argument)
//!   - If length = 0, returns empty string
//!
//! ## Return Value
//!
//! Returns a Variant (String):
//! - Substring of specified length starting at start position
//! - Uses 1-based indexing (first character is 1, not 0)
//! - Returns empty string if start > string length
//! - Returns remaining characters if length extends past end of string
//! - Returns Null if input string is Null
//! - Returns empty string if input string is empty
//! - Returns empty string if length = 0
//!
//! ## Remarks
//!
//! The Mid function extracts a substring:
//!
//! - **1-based indexing**: First character is at position 1 (not 0 like in many languages)
//! - **Optional length**: If omitted, returns from start to end of string
//! - **Bounds handling**: If start or length exceed string bounds, adjusts gracefully
//! - **No error on overflow**: Returns available characters without error
//! - **Null propagation**: Returns Null if input is Null
//! - **Common use**: Extract portions of strings, parse data, substring operations
//! - **Related statement**: Mid statement assigns to substring (Mid(s, 1, 3) = "abc")
//! - **Similar to**: Left (from start), Right (from end), `InStr` (find position)
//! - **Performance**: Fast operation, optimized in VB6
//! - **String immutability**: Returns new string, does not modify original
//! - **Unicode support**: Works with Unicode strings in VB6
//! - **Type conversion**: Automatically converts numeric strings
//! - **Available in**: All VB versions, VBA, `VBScript`
//!
//! ## Typical Uses
//!
//! 1. **Extract Substring**
//!    ```vb
//!    middle = Mid("Hello World", 7, 5)  ' Returns "World"
//!    ```
//!
//! 2. **Parse Fixed-Width Data**
//!    ```vb
//!    customerID = Mid(record, 1, 10)
//!    customerName = Mid(record, 11, 30)
//!    ```
//!
//! 3. **Extract from Position to End**
//!    ```vb
//!    remainder = Mid(text, 5)  ' From position 5 to end
//!    ```
//!
//! 4. **Parse Delimited Data**
//!    ```vb
//!    pos = InStr(data, ",")
//!    firstField = Mid(data, 1, pos - 1)
//!    ```
//!
//! 5. **Skip Characters**
//!    ```vb
//!    withoutPrefix = Mid(text, 4)  ' Skip first 3 characters
//!    ```
//!
//! 6. **Extract Single Character**
//!    ```vb
//!    char = Mid(text, i, 1)  ' Get character at position i
//!    ```
//!
//! 7. **Data Validation**
//!    ```vb
//!    areaCode = Mid(phoneNumber, 2, 3)  ' Extract area code
//!    ```
//!
//! 8. **String Manipulation**
//!    ```vb
//!    modified = Left(s, 5) & "***" & Mid(s, 9)  ' Mask middle
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Usage
//! ```vb
//! Dim result As String
//! Dim text As String
//!
//! text = "Hello World"
//!
//! ' Extract with length
//! result = Mid(text, 1, 5)   ' Returns "Hello"
//! result = Mid(text, 7, 5)   ' Returns "World"
//! result = Mid(text, 3, 3)   ' Returns "llo"
//!
//! ' Extract to end (no length parameter)
//! result = Mid(text, 7)      ' Returns "World"
//! result = Mid(text, 1)      ' Returns "Hello World"
//!
//! ' Edge cases
//! result = Mid(text, 20)     ' Returns "" (start past end)
//! result = Mid(text, 7, 100) ' Returns "World" (length past end)
//! result = Mid(text, 5, 0)   ' Returns "" (zero length)
//! ```
//!
//! ### Example 2: Parse Fixed-Width Record
//! ```vb
//! Sub ParseFixedWidthRecord()
//!     Dim record As String
//!     Dim customerID As String
//!     Dim customerName As String
//!     Dim city As String
//!     Dim state As String
//!     Dim zipCode As String
//!     
//!     ' Example: "CUST001   John Smith            New York    NY12345"
//!     record = "CUST001   John Smith            New York    NY12345"
//!     
//!     ' Parse fixed-width fields
//!     customerID = RTrim(Mid(record, 1, 10))     ' Positions 1-10
//!     customerName = RTrim(Mid(record, 11, 22))  ' Positions 11-32
//!     city = RTrim(Mid(record, 33, 12))          ' Positions 33-44
//!     state = Mid(record, 45, 2)                  ' Positions 45-46
//!     zipCode = Mid(record, 47, 5)                ' Positions 47-51
//!     
//!     Debug.Print "ID: " & customerID
//!     Debug.Print "Name: " & customerName
//!     Debug.Print "City: " & city
//!     Debug.Print "State: " & state
//!     Debug.Print "Zip: " & zipCode
//! End Sub
//! ```
//!
//! ### Example 3: Extract File Extension
//! ```vb
//! Function GetFileExtension(ByVal filename As String) As String
//!     Dim dotPos As Long
//!     
//!     ' Find last dot
//!     dotPos = InStrRev(filename, ".")
//!     
//!     If dotPos > 0 Then
//!         ' Extract extension (without dot)
//!         GetFileExtension = Mid(filename, dotPos + 1)
//!     Else
//!         GetFileExtension = ""
//!     End If
//! End Function
//!
//! ' Usage:
//! ' ext = GetFileExtension("document.txt")      ' Returns "txt"
//! ' ext = GetFileExtension("photo.jpg")         ' Returns "jpg"
//! ' ext = GetFileExtension("archive.tar.gz")    ' Returns "gz"
//! ' ext = GetFileExtension("README")            ' Returns ""
//! ```
//!
//! ### Example 4: Parse Delimited String
//! ```vb
//! Sub ParseCSVLine(ByVal line As String)
//!     Dim pos1 As Long, pos2 As Long
//!     Dim field1 As String, field2 As String, field3 As String
//!     
//!     ' Parse: "Smith,John,123 Main St"
//!     
//!     ' Find first comma
//!     pos1 = InStr(line, ",")
//!     If pos1 > 0 Then
//!         field1 = Mid(line, 1, pos1 - 1)
//!         
//!         ' Find second comma
//!         pos2 = InStr(pos1 + 1, line, ",")
//!         If pos2 > 0 Then
//!             field2 = Mid(line, pos1 + 1, pos2 - pos1 - 1)
//!             field3 = Mid(line, pos2 + 1)  ' Rest of string
//!         Else
//!             field2 = Mid(line, pos1 + 1)  ' No third field
//!             field3 = ""
//!         End If
//!     Else
//!         field1 = line
//!         field2 = ""
//!         field3 = ""
//!     End If
//!     
//!     Debug.Print "Last Name: " & field1
//!     Debug.Print "First Name: " & field2
//!     Debug.Print "Address: " & field3
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `SafeMid` (handle Null)
//! ```vb
//! Function SafeMid(ByVal text As Variant, _
//!                  ByVal start As Long, _
//!                  Optional ByVal length As Long = -1) As String
//!     If IsNull(text) Then
//!         SafeMid = ""
//!     ElseIf length = -1 Then
//!         SafeMid = Mid(text, start)
//!     Else
//!         SafeMid = Mid(text, start, length)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: `GetChar` (extract single character)
//! ```vb
//! Function GetChar(ByVal text As String, ByVal position As Long) As String
//!     If position >= 1 And position <= Len(text) Then
//!         GetChar = Mid(text, position, 1)
//!     Else
//!         GetChar = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `SkipChars` (remove prefix)
//! ```vb
//! Function SkipChars(ByVal text As String, ByVal count As Long) As String
//!     If count >= Len(text) Then
//!         SkipChars = ""
//!     ElseIf count <= 0 Then
//!         SkipChars = text
//!     Else
//!         SkipChars = Mid(text, count + 1)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `ExtractBetween`
//! ```vb
//! Function ExtractBetween(ByVal text As String, _
//!                         ByVal startPos As Long, _
//!                         ByVal endPos As Long) As String
//!     If endPos >= startPos And startPos >= 1 Then
//!         ExtractBetween = Mid(text, startPos, endPos - startPos + 1)
//!     Else
//!         ExtractBetween = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `MaskMiddle` (hide sensitive data)
//! ```vb
//! Function MaskMiddle(ByVal text As String, _
//!                     ByVal visibleStart As Long, _
//!                     ByVal visibleEnd As Long) As String
//!     Dim textLen As Long
//!     textLen = Len(text)
//!     
//!     If textLen <= visibleStart + visibleEnd Then
//!         MaskMiddle = text
//!     Else
//!         MaskMiddle = Left(text, visibleStart) & _
//!                     String(textLen - visibleStart - visibleEnd, "*") & _
//!                     Right(text, visibleEnd)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 6: `ParseFixedField`
//! ```vb
//! Function ParseFixedField(ByVal record As String, _
//!                         ByVal startPos As Long, _
//!                         ByVal fieldWidth As Long) As String
//!     ParseFixedField = RTrim(Mid(record, startPos, fieldWidth))
//! End Function
//! ```
//!
//! ### Pattern 7: `GetSubstringAfter`
//! ```vb
//! Function GetSubstringAfter(ByVal text As String, _
//!                           ByVal delimiter As String) As String
//!     Dim pos As Long
//!     pos = InStr(text, delimiter)
//!     
//!     If pos > 0 Then
//!         GetSubstringAfter = Mid(text, pos + Len(delimiter))
//!     Else
//!         GetSubstringAfter = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: `GetSubstringBefore`
//! ```vb
//! Function GetSubstringBefore(ByVal text As String, _
//!                            ByVal delimiter As String) As String
//!     Dim pos As Long
//!     pos = InStr(text, delimiter)
//!     
//!     If pos > 0 Then
//!         GetSubstringBefore = Mid(text, 1, pos - 1)
//!     Else
//!         GetSubstringBefore = text
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `ReplaceMiddle`
//! ```vb
//! Function ReplaceMiddle(ByVal text As String, _
//!                       ByVal start As Long, _
//!                       ByVal length As Long, _
//!                       ByVal replacement As String) As String
//!     ReplaceMiddle = Left(text, start - 1) & _
//!                    replacement & _
//!                    Mid(text, start + length)
//! End Function
//! ```
//!
//! ### Pattern 10: `ExtractWord`
//! ```vb
//! Function ExtractWord(ByVal text As String, ByVal wordNum As Long) As String
//!     Dim words() As String
//!     words = Split(Trim(text))
//!     
//!     If wordNum >= 1 And wordNum <= UBound(words) + 1 Then
//!         ExtractWord = words(wordNum - 1)
//!     Else
//!         ExtractWord = ""
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Fixed-Width File Parser
//! ```vb
//! ' Class: FixedWidthParser
//! Private Type FieldDefinition
//!     Name As String
//!     StartPos As Long
//!     Length As Long
//!     TrimSpaces As Boolean
//! End Type
//!
//! Private m_fields() As FieldDefinition
//! Private m_fieldCount As Long
//!
//! Public Sub AddField(ByVal name As String, _
//!                     ByVal startPos As Long, _
//!                     ByVal length As Long, _
//!                     Optional ByVal trimSpaces As Boolean = True)
//!     m_fieldCount = m_fieldCount + 1
//!     ReDim Preserve m_fields(1 To m_fieldCount)
//!     
//!     With m_fields(m_fieldCount)
//!         .Name = name
//!         .StartPos = startPos
//!         .Length = length
//!         .TrimSpaces = trimSpaces
//!     End With
//! End Sub
//!
//! Public Function ParseRecord(ByVal record As String) As Collection
//!     Dim result As New Collection
//!     Dim i As Long
//!     Dim value As String
//!     
//!     For i = 1 To m_fieldCount
//!         With m_fields(i)
//!             value = Mid(record, .StartPos, .Length)
//!             
//!             If .TrimSpaces Then
//!                 value = Trim(value)
//!             End If
//!             
//!             result.Add value, .Name
//!         End With
//!     Next i
//!     
//!     Set ParseRecord = result
//! End Function
//!
//! Public Sub ParseFile(ByVal filename As String, _
//!                     ByVal outputCollection As Collection)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim recordData As Collection
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         Set recordData = ParseRecord(line)
//!         outputCollection.Add recordData
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Example 2: String Tokenizer
//! ```vb
//! ' Class: StringTokenizer
//! Private m_text As String
//! Private m_position As Long
//! Private m_length As Long
//!
//! Public Sub Initialize(ByVal text As String)
//!     m_text = text
//!     m_position = 1
//!     m_length = Len(text)
//! End Sub
//!
//! Public Function HasMoreTokens() As Boolean
//!     HasMoreTokens = (m_position <= m_length)
//! End Function
//!
//! Public Function NextToken(ByVal delimiter As String) As String
//!     Dim delimPos As Long
//!     Dim token As String
//!     
//!     If m_position > m_length Then
//!         NextToken = ""
//!         Exit Function
//!     End If
//!     
//!     ' Find next delimiter
//!     delimPos = InStr(m_position, m_text, delimiter)
//!     
//!     If delimPos = 0 Then
//!         ' No more delimiters, return rest of string
//!         token = Mid(m_text, m_position)
//!         m_position = m_length + 1
//!     Else
//!         ' Extract token
//!         token = Mid(m_text, m_position, delimPos - m_position)
//!         m_position = delimPos + Len(delimiter)
//!     End If
//!     
//!     NextToken = token
//! End Function
//!
//! Public Function PeekToken(ByVal delimiter As String) As String
//!     Dim savedPos As Long
//!     savedPos = m_position
//!     PeekToken = NextToken(delimiter)
//!     m_position = savedPos
//! End Function
//!
//! Public Sub Reset()
//!     m_position = 1
//! End Sub
//!
//! Public Property Get Position() As Long
//!     Position = m_position
//! End Property
//!
//! Public Property Let Position(ByVal newPos As Long)
//!     If newPos >= 1 And newPos <= m_length + 1 Then
//!         m_position = newPos
//!     End If
//! End Property
//! ```
//!
//! ### Example 3: String Masking Utility
//! ```vb
//! ' Module: StringMasking
//!
//! Public Function MaskCreditCard(ByVal cardNumber As String) As String
//!     ' Show last 4 digits: **** **** **** 1234
//!     Dim cleaned As String
//!     cleaned = Replace(cardNumber, " ", "")
//!     cleaned = Replace(cleaned, "-", "")
//!     
//!     If Len(cleaned) >= 4 Then
//!         MaskCreditCard = String(Len(cleaned) - 4, "*") & Mid(cleaned, Len(cleaned) - 3)
//!     Else
//!         MaskCreditCard = String(Len(cleaned), "*")
//!     End If
//! End Function
//!
//! Public Function MaskSSN(ByVal ssn As String) As String
//!     ' Show last 4 digits: ***-**-1234
//!     Dim cleaned As String
//!     cleaned = Replace(ssn, "-", "")
//!     
//!     If Len(cleaned) = 9 Then
//!         MaskSSN = "***-**-" & Mid(cleaned, 6)
//!     Else
//!         MaskSSN = String(Len(ssn), "*")
//!     End If
//! End Function
//!
//! Public Function MaskEmail(ByVal email As String) As String
//!     ' Show first 2 chars and domain: jo****@example.com
//!     Dim atPos As Long
//!     Dim localPart As String
//!     Dim domainPart As String
//!     
//!     atPos = InStr(email, "@")
//!     
//!     If atPos > 2 Then
//!         localPart = Mid(email, 1, 2) & String(atPos - 3, "*")
//!         domainPart = Mid(email, atPos)
//!         MaskEmail = localPart & domainPart
//!     Else
//!         MaskEmail = String(Len(email), "*")
//!     End If
//! End Function
//!
//! Public Function MaskPhone(ByVal phoneNumber As String) As String
//!     ' Show area code and last 2: (123) ***-**34
//!     Dim cleaned As String
//!     cleaned = Replace(phoneNumber, "(", "")
//!     cleaned = Replace(cleaned, ")", "")
//!     cleaned = Replace(cleaned, " ", "")
//!     cleaned = Replace(cleaned, "-", "")
//!     
//!     If Len(cleaned) = 10 Then
//!         MaskPhone = "(" & Mid(cleaned, 1, 3) & ") ***-**" & Mid(cleaned, 9)
//!     Else
//!         MaskPhone = String(Len(phoneNumber), "*")
//!     End If
//! End Function
//! ```
//!
//! ### Example 4: CSV Parser with Quoted Fields
//! ```vb
//! ' Module: CSVParser
//!
//! Public Function ParseCSVLine(ByVal line As String) As Variant
//!     Dim fields() As String
//!     Dim fieldCount As Long
//!     Dim position As Long
//!     Dim length As Long
//!     Dim inQuote As Boolean
//!     Dim currentField As String
//!     Dim ch As String
//!     
//!     length = Len(line)
//!     position = 1
//!     fieldCount = 0
//!     inQuote = False
//!     currentField = ""
//!     
//!     Do While position <= length
//!         ch = Mid(line, position, 1)
//!         
//!         Select Case ch
//!             Case """"  ' Quote
//!                 If inQuote Then
//!                     ' Check for escaped quote
//!                     If position < length Then
//!                         If Mid(line, position + 1, 1) = """" Then
//!                             currentField = currentField & """"
//!                             position = position + 1
//!                         Else
//!                             inQuote = False
//!                         End If
//!                     Else
//!                         inQuote = False
//!                     End If
//!                 Else
//!                     inQuote = True
//!                 End If
//!                 
//!             Case ","
//!                 If inQuote Then
//!                     currentField = currentField & ch
//!                 Else
//!                     ' End of field
//!                     fieldCount = fieldCount + 1
//!                     ReDim Preserve fields(1 To fieldCount)
//!                     fields(fieldCount) = currentField
//!                     currentField = ""
//!                 End If
//!                 
//!             Case Else
//!                 currentField = currentField & ch
//!         End Select
//!         
//!         position = position + 1
//!     Loop
//!     
//!     ' Add last field
//!     fieldCount = fieldCount + 1
//!     ReDim Preserve fields(1 To fieldCount)
//!     fields(fieldCount) = currentField
//!     
//!     ParseCSVLine = fields
//! End Function
//!
//! Public Function GetCSVField(ByVal line As String, _
//!                            ByVal fieldIndex As Long) As String
//!     Dim fields As Variant
//!     fields = ParseCSVLine(line)
//!     
//!     If fieldIndex >= LBound(fields) And fieldIndex <= UBound(fields) Then
//!         GetCSVField = fields(fieldIndex)
//!     Else
//!         GetCSVField = ""
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 5: Invalid procedure call or argument
//! ' - start < 1
//! ' - length < 0
//!
//! ' Safe extraction with error handling
//! Function SafeExtract(ByVal text As String, _
//!                     ByVal start As Long, _
//!                     ByVal length As Long) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(text) Then
//!         SafeExtract = ""
//!     ElseIf start < 1 Or length < 0 Then
//!         SafeExtract = ""
//!     ElseIf start > Len(text) Then
//!         SafeExtract = ""
//!     Else
//!         SafeExtract = Mid(text, start, length)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeExtract = ""
//! End Function
//!
//! ' Handle Null values
//! Function MidSafe(ByVal text As Variant, _
//!                  ByVal start As Long, _
//!                  Optional ByVal length As Variant = Empty) As String
//!     If IsNull(text) Then
//!         MidSafe = ""
//!     ElseIf IsEmpty(length) Then
//!         MidSafe = Mid(text, start)
//!     Else
//!         MidSafe = Mid(text, start, length)
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: String extraction is highly optimized in VB6
//! - **Creates New String**: Does not modify original (immutable)
//! - **Avoid in Tight Loops**: Cache result if using multiple times
//! - **Better than**: Repeated Left/Right operations for complex parsing
//! - **Consider Split**: For delimited data, Split may be faster
//! - **String Builder**: For concatenating many Mid results, use `StringBuilder` pattern
//!
//! ## Best Practices
//!
//! 1. **Remember 1-based indexing** - First character is position 1, not 0
//! 2. **Validate inputs** - Check start and length before calling Mid
//! 3. **Handle Null gracefully** - Use `IsNull` check for Variant inputs
//! 4. **Omit length when extracting to end** - More readable: Mid(s, 5) vs Mid(s, 5, Len(s)-4)
//! 5. **Combine with Trim** - Clean whitespace from extracted fields
//! 6. **Use with `InStr`** - Find position, then extract with Mid
//! 7. **Cache Len results** - If using Len(string) multiple times
//! 8. **Document field positions** - For fixed-width parsing, use constants
//! 9. **Test edge cases** - Empty strings, start past end, Null values
//! 10. **Consider alternatives** - Split for delimited data, Left/Right for ends
//!
//! ## Comparison with Related Functions
//!
//! | Function | Extracts From | Parameters | Use Case |
//! |----------|--------------|------------|----------|
//! | **Mid** | Any position | start, [length] | General substring extraction |
//! | **Left** | Beginning | length | Get first N characters |
//! | **Right** | End | length | Get last N characters |
//! | **`InStr`** | N/A (finds) | [start,] string1, string2 | Find position of substring |
//!
//! ## Mid vs Left vs Right
//!
//! ```vb
//! Dim text As String
//! text = "Hello World"
//!
//! ' Mid - extract from any position
//! Debug.Print Mid(text, 7, 5)    ' "World" (position 7, length 5)
//! Debug.Print Mid(text, 3, 3)    ' "llo" (position 3, length 3)
//! Debug.Print Mid(text, 7)       ' "World" (from position 7 to end)
//!
//! ' Left - extract from beginning
//! Debug.Print Left(text, 5)      ' "Hello" (first 5 characters)
//!
//! ' Right - extract from end
//! Debug.Print Right(text, 5)     ' "World" (last 5 characters)
//!
//! ' Equivalent operations
//! Debug.Print Mid(text, 1, 5)    ' "Hello" (same as Left(text, 5))
//! Debug.Print Mid(text, 7)       ' "World" (same as Right(text, 5))
//! ```
//!
//! ## Mid Function vs Mid Statement
//!
//! ```vb
//! Dim text As String
//! text = "Hello World"
//!
//! ' Mid Function - returns substring (does not modify)
//! Dim result As String
//! result = Mid(text, 1, 5)       ' result = "Hello", text unchanged
//!
//! ' Mid Statement - modifies string in place
//! Mid(text, 1, 5) = "Goodbye"    ' text = "GoodbyWorld" (replaces 5 chars)
//! Mid(text, 7, 5) = "Earth"      ' text = "Hello Earth" (replaces 5 chars)
//!
//! ' Note: Mid statement exists but is less commonly used
//! ' Mid function is much more common for extraction
//! ```
//!
//! ## 1-Based vs 0-Based Indexing
//!
//! ```vb
//! ' VB6 Mid uses 1-based indexing
//! Dim text As String
//! text = "ABCDE"
//!
//! Debug.Print Mid(text, 1, 1)    ' "A" (first character is position 1)
//! Debug.Print Mid(text, 2, 1)    ' "B" (second character is position 2)
//! Debug.Print Mid(text, 5, 1)    ' "E" (fifth character is position 5)
//!
//! ' Compare to 0-based languages (JavaScript, C#, etc.)
//! ' text[0]   // "A" (first character is index 0)
//! ' text[1]   // "B" (second character is index 1)
//! ' text[4]   // "E" (fifth character is index 4)
//!
//! ' When converting between languages:
//! ' VB6: Mid(text, n, 1)
//! ' C#:  text[n-1]
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Available in `VBScript`
//! - Works with ANSI and Unicode strings
//! - **1-based indexing** (first character is 1)
//! - Returns new string (original unchanged)
//! - Handles Null by returning Null
//! - No error if start or length exceed bounds (adjusts gracefully)
//! - Same behavior across all Windows versions
//!
//! ## Limitations
//!
//! - **1-based indexing**: Not 0-based like most modern languages
//! - **Creates new string**: Cannot modify string in place (use Mid statement for that)
//! - **No negative indices**: Cannot count from end like Python
//! - **No regex support**: For pattern matching, use `RegExp` object
//! - **Error on invalid start**: start < 1 causes error 5
//! - **Error on negative length**: length < 0 causes error 5
//!
//! ## Related Functions
//!
//! - `Left`: Returns leftmost characters from string
//! - `Right`: Returns rightmost characters from string
//! - `InStr`: Finds position of substring
//! - `InStrRev`: Finds position of substring from end
//! - `Len`: Returns string length
//! - `LTrim`/`RTrim`/`Trim`: Removes spaces
//! - `Split`: Splits string into array
//! - `Replace`: Replaces substring occurrences
//! - `Mid` statement: Replaces characters within string

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn mid_basic() {
        let source = r#"
            result = Mid("Hello World", 7, 5)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_two_arguments() {
        let source = r#"
            result = Mid(text, 5)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_variable() {
        let source = r#"
            substring = Mid(fullText, startPos, length)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_if_statement() {
        let source = r#"
            If Mid(text, 1, 5) = "Hello" Then
                MsgBox "Match"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
    }

    #[test]
    fn mid_function_return() {
        let source = r#"
            Function GetSubstring() As String
                GetSubstring = Mid(data, 10, 20)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_with_instr() {
        let source = r#"
            pos = InStr(text, ",")
            field = Mid(text, 1, pos - 1)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_debug_print() {
        let source = r#"
            Debug.Print Mid(message, 5, 10)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_with_statement() {
        let source = r#"
            With record
                .ID = Mid(.FullData, 1, 10)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_select_case() {
        let source = r#"
            Select Case Mid(code, 1, 2)
                Case "AA"
                    MsgBox "Type A"
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_elseif() {
        let source = r#"
            If code = "" Then
                status = "Empty"
            ElseIf Mid(code, 1, 1) = "X" Then
                status = "Special"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_parentheses() {
        let source = r#"
            result = (Mid(text, 3, 5))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_iif() {
        let source = r#"
            result = IIf(Mid(text, 1, 1) = "A", "Yes", "No")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_in_class() {
        let source = r#"
            Private Sub ExtractData()
                m_code = Mid(m_rawData, 1, 5)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_function_argument() {
        let source = r#"
            Call ProcessText(Mid(input, 10, 50))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_property_assignment() {
        let source = r#"
            MyObject.Substring = Mid(fullString, 5, 10)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_array_assignment() {
        let source = r#"
            fields(i) = Mid(record, pos, fieldWidth)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_for_loop() {
        let source = r#"
            For i = 1 To Len(text)
                char = Mid(text, i, 1)
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_while_wend() {
        let source = r#"
            While pos <= Len(data)
                field = Mid(data, pos, 10)
                pos = pos + 10
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_do_while() {
        let source = r#"
            Do While i <= recordCount
                customerID = Mid(records(i), 1, 10)
                i = i + 1
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_do_until() {
        let source = r#"
            Do Until pos > Len(text)
                token = Mid(text, pos, delimPos - pos)
                pos = delimPos + 1
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_msgbox() {
        let source = r#"
            MsgBox Mid(errorMessage, 1, 50)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_concatenation() {
        let source = r#"
            fullName = Mid(firstName, 1, 1) & ". " & lastName
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_comparison() {
        let source = r#"
            If Mid(text1, 1, 5) = Mid(text2, 1, 5) Then
                MsgBox "Match"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_with_trim() {
        let source = r#"
            cleanField = Trim(Mid(record, 10, 30))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_fixed_width() {
        let source = r#"
            customerID = Mid(record, 1, 10)
            customerName = Mid(record, 11, 30)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_nested() {
        let source = r#"
            result = Mid(Mid(text, 5, 20), 3, 10)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn mid_single_char() {
        let source = r#"
            char = Mid(text, position, 1)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Mid"));
        assert!(text.contains("Identifier"));
    }
}
