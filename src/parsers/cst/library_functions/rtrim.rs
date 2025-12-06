//! # `RTrim` Function
//!
//! Returns a String containing a copy of a specified string with trailing spaces removed.
//!
//! ## Syntax
//!
//! ```vb
//! RTrim(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which trailing spaces are to be removed
//!   - Can be any valid string expression
//!   - If string is Null, returns Null
//!   - Empty string returns empty string
//!
//! ## Return Value
//!
//! Returns a String (or Variant):
//! - Copy of string with trailing spaces removed
//! - Removes only spaces (ASCII 32) from the right
//! - Does not remove tabs, newlines, or other whitespace characters
//! - Returns Null if input is Null
//! - Returns empty string if input is empty or all spaces
//! - Leading spaces are preserved
//! - Internal spaces are preserved
//!
//! ## Remarks
//!
//! The `RTrim` function removes trailing spaces:
//!
//! - Removes only space characters (ASCII 32) from the right side
//! - Does not remove tabs (Chr(9)), line feeds (Chr(10)), or carriage returns (Chr(13))
//! - Does not remove non-breaking spaces or other Unicode whitespace
//! - Leading spaces are not affected
//! - Internal spaces between words are preserved
//! - Often used to clean up output formatting
//! - Commonly paired with `LTrim` or used with Trim
//! - Null input returns Null (propagates Null)
//! - Empty string input returns empty string
//! - String of only spaces returns empty string
//! - Does not modify the original string (returns new string)
//! - Can be used with Variant variables
//! - Common in report generation and text alignment
//! - Used to remove padding from fixed-width fields
//! - Essential for cleaning exported data
//! - Part of the VB6 string manipulation library
//! - Available in all VB versions
//! - Related to `LTrim` (removes leading spaces) and Trim (removes both)
//!
//! ## Typical Uses
//!
//! 1. **Remove Trailing Spaces**
//!    ```vb
//!    cleanText = RTrim("Hello   ")
//!    ```
//!
//! 2. **Clean Fixed-Width Output**
//!    ```vb
//!    outputLine = RTrim(paddedField)
//!    ```
//!
//! 3. **Format Display**
//!    ```vb
//!    lblName.Caption = RTrim(recordset("Name"))
//!    ```
//!
//! 4. **Process Data Export**
//!    ```vb
//!    exportValue = RTrim(databaseField)
//!    ```
//!
//! 5. **Normalize Text**
//!    ```vb
//!    normalizedText = LTrim(RTrim(inputText))
//!    ```
//!
//! 6. **String Comparison**
//!    ```vb
//!    If RTrim(text1) = RTrim(text2) Then
//!        ' Equal when ignoring trailing spaces
//!    End If
//!    ```
//!
//! 7. **Write to File**
//!    ```vb
//!    Print #1, RTrim(record.Field1)
//!    ```
//!
//! 8. **Report Generation**
//!    ```vb
//!    reportLine = RTrim(customerName) & " - " & orderID
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Usage
//! ```vb
//! Dim result As String
//!
//! result = RTrim("Hello   ")           ' Returns "Hello"
//! result = RTrim("   Hello")           ' Returns "   Hello" (leading preserved)
//! result = RTrim("   Hello World   ")  ' Returns "   Hello World"
//! result = RTrim("NoSpaces")           ' Returns "NoSpaces"
//! result = RTrim("     ")              ' Returns ""
//! result = RTrim("")                   ' Returns ""
//! ```
//!
//! ### Example 2: Fixed-Width File Export
//! ```vb
//! Sub ExportToFixedWidth(ByVal filename As String)
//!     Dim rs As ADODB.Recordset
//!     Dim fileNum As Integer
//!     Dim line As String
//!     
//!     Set rs = New ADODB.Recordset
//!     rs.Open "SELECT * FROM Customers", conn
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     Do While Not rs.EOF
//!         ' Build fixed-width line and remove trailing spaces
//!         line = Left(rs("CustomerID") & Space(10), 10) & _
//!                Left(rs("CompanyName") & Space(40), 40) & _
//!                Left(rs("City") & Space(20), 20)
//!         
//!         ' Write without trailing spaces
//!         Print #fileNum, RTrim(line)
//!         
//!         rs.MoveNext
//!     Loop
//!     
//!     Close #fileNum
//!     rs.Close
//!     Set rs = Nothing
//! End Sub
//! ```
//!
//! ### Example 3: Database Field Cleanup
//! ```vb
//! Function GetCustomerName(ByVal customerID As String) As String
//!     Dim rs As ADODB.Recordset
//!     Dim sql As String
//!     
//!     sql = "SELECT CustomerName FROM Customers WHERE CustomerID = '" & customerID & "'"
//!     
//!     Set rs = New ADODB.Recordset
//!     rs.Open sql, conn
//!     
//!     If Not rs.EOF Then
//!         ' Remove trailing spaces from database field (may be CHAR type)
//!         GetCustomerName = RTrim(rs("CustomerName") & "")
//!     Else
//!         GetCustomerName = ""
//!     End If
//!     
//!     rs.Close
//!     Set rs = Nothing
//! End Function
//! ```
//!
//! ### Example 4: Report Formatting
//! ```vb
//! Sub GenerateReport()
//!     Dim rs As ADODB.Recordset
//!     Dim reportLine As String
//!     
//!     Set rs = New ADODB.Recordset
//!     rs.Open "SELECT * FROM Orders", conn
//!     
//!     lstReport.Clear
//!     
//!     Do While Not rs.EOF
//!         ' Format report line with proper spacing
//!         reportLine = RTrim(rs("CustomerName") & "") & " - " & _
//!                     "Order #" & rs("OrderID") & " - " & _
//!                     Format(rs("OrderDate"), "mm/dd/yyyy")
//!         
//!         lstReport.AddItem reportLine
//!         
//!         rs.MoveNext
//!     Loop
//!     
//!     rs.Close
//!     Set rs = Nothing
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `FullTrim` (combine with `LTrim`)
//! ```vb
//! Function FullTrim(ByVal text As String) As String
//!     FullTrim = LTrim(RTrim(text))
//!     ' Note: Can also use built-in Trim() function
//! End Function
//! ```
//!
//! ### Pattern 2: `SafeRTrim` (handle Null)
//! ```vb
//! Function SafeRTrim(ByVal text As Variant) As String
//!     If IsNull(text) Then
//!         SafeRTrim = ""
//!     Else
//!         SafeRTrim = RTrim(text)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `CleanDatabaseField`
//! ```vb
//! Function CleanDatabaseField(ByVal rs As Recordset, _
//!                            ByVal fieldName As String) As String
//!     If Not IsNull(rs(fieldName)) Then
//!         CleanDatabaseField = RTrim(rs(fieldName) & "")
//!     Else
//!         CleanDatabaseField = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `TrimTrailingSpaces`
//! ```vb
//! Sub TrimTrailingSpaces(fields() As String)
//!     Dim i As Integer
//!     For i = LBound(fields) To UBound(fields)
//!         fields(i) = RTrim(fields(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 5: `CompareIgnoreTrailing`
//! ```vb
//! Function CompareIgnoreTrailing(ByVal str1 As String, _
//!                                ByVal str2 As String) As Boolean
//!     CompareIgnoreTrailing = (RTrim(str1) = RTrim(str2))
//! End Function
//! ```
//!
//! ### Pattern 6: `FormatFixedWidth`
//! ```vb
//! Function FormatFixedWidth(ByVal text As String, _
//!                          ByVal width As Integer) As String
//!     Dim padded As String
//!     padded = Left(text & Space(width), width)
//!     FormatFixedWidth = RTrim(padded)
//! End Function
//! ```
//!
//! ### Pattern 7: `CleanRecordsetField`
//! ```vb
//! Function CleanRecordsetField(ByVal rs As Recordset, _
//!                             ByVal fieldName As String, _
//!                             Optional ByVal defaultValue As String = "") As String
//!     On Error Resume Next
//!     If Not IsNull(rs(fieldName)) Then
//!         CleanRecordsetField = RTrim(CStr(rs(fieldName)))
//!     Else
//!         CleanRecordsetField = defaultValue
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: `ExportCleanLine`
//! ```vb
//! Sub ExportCleanLine(ByVal fileNum As Integer, _
//!                     ByVal text As String)
//!     Print #fileNum, RTrim(text)
//! End Sub
//! ```
//!
//! ### Pattern 9: `RemoveTrailingPadding`
//! ```vb
//! Function RemoveTrailingPadding(ByVal paddedText As String) As String
//!     RemoveTrailingPadding = RTrim(paddedText)
//! End Function
//! ```
//!
//! ### Pattern 10: `CleanDisplayText`
//! ```vb
//! Sub CleanDisplayText(ByVal ctrl As Control)
//!     If TypeOf ctrl Is Label Or TypeOf ctrl Is TextBox Then
//!         ctrl.Caption = RTrim(ctrl.Caption)
//!     End If
//! End Sub
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Fixed-Width File Processor
//! ```vb
//! ' Class: FixedWidthExporter
//! Private m_fileNum As Integer
//! Private m_isOpen As Boolean
//!
//! Public Sub OpenFile(ByVal filename As String)
//!     If m_isOpen Then CloseFile
//!     
//!     m_fileNum = FreeFile
//!     Open filename For Output As #m_fileNum
//!     m_isOpen = True
//! End Sub
//!
//! Public Sub WriteField(ByVal text As String, ByVal width As Integer)
//!     Dim padded As String
//!     
//!     ' Pad to width and remove trailing spaces
//!     padded = Left(text & Space(width), width)
//!     Print #m_fileNum, RTrim(padded);
//! End Sub
//!
//! Public Sub WriteLine()
//!     Print #m_fileNum, ""
//! End Sub
//!
//! Public Sub WriteRecord(ParamArray fields() As Variant)
//!     Dim i As Integer
//!     Dim line As String
//!     
//!     For i = LBound(fields) To UBound(fields)
//!         If i Mod 2 = 0 Then
//!             ' Even index = text
//!             line = line & Left(fields(i) & Space(fields(i + 1)), fields(i + 1))
//!         End If
//!     Next i
//!     
//!     Print #m_fileNum, RTrim(line)
//! End Sub
//!
//! Public Sub CloseFile()
//!     If m_isOpen Then
//!         Close #m_fileNum
//!         m_isOpen = False
//!     End If
//! End Sub
//!
//! Private Sub Class_Terminate()
//!     CloseFile
//! End Sub
//! ```
//!
//! ### Example 2: Database Field Cleaner
//! ```vb
//! ' Class: RecordsetCleaner
//! Private m_trimLeading As Boolean
//! Private m_trimTrailing As Boolean
//!
//! Private Sub Class_Initialize()
//!     m_trimLeading = False
//!     m_trimTrailing = True
//! End Sub
//!
//! Public Property Let TrimLeading(ByVal value As Boolean)
//!     m_trimLeading = value
//! End Property
//!
//! Public Property Let TrimTrailing(ByVal value As Boolean)
//!     m_trimTrailing = value
//! End Property
//!
//! Public Function GetCleanField(ByVal rs As Recordset, _
//!                               ByVal fieldName As String) As String
//!     Dim value As String
//!     
//!     If IsNull(rs(fieldName)) Then
//!         GetCleanField = ""
//!         Exit Function
//!     End If
//!     
//!     value = CStr(rs(fieldName))
//!     
//!     If m_trimTrailing Then
//!         value = RTrim(value)
//!     End If
//!     
//!     If m_trimLeading Then
//!         value = LTrim(value)
//!     End If
//!     
//!     GetCleanField = value
//! End Function
//!
//! Public Sub CleanRecordset(ByVal rs As Recordset)
//!     Dim fld As Field
//!     
//!     If rs.EOF And rs.BOF Then Exit Sub
//!     
//!     rs.MoveFirst
//!     Do While Not rs.EOF
//!         For Each fld In rs.Fields
//!             If fld.Type = adVarChar Or fld.Type = adChar Then
//!                 If Not IsNull(fld.Value) Then
//!                     fld.Value = GetCleanField(rs, fld.Name)
//!                 End If
//!             End If
//!         Next fld
//!         rs.MoveNext
//!     Loop
//! End Sub
//! ```
//!
//! ### Example 3: Report Generator
//! ```vb
//! ' Class: ReportGenerator
//! Private m_lines As Collection
//! Private m_columnWidths() As Integer
//!
//! Public Sub Initialize(columnWidths() As Integer)
//!     Set m_lines = New Collection
//!     m_columnWidths = columnWidths
//! End Sub
//!
//! Public Sub AddRow(ParamArray values() As Variant)
//!     Dim i As Integer
//!     Dim line As String
//!     Dim cellValue As String
//!     
//!     For i = LBound(values) To UBound(values)
//!         If i <= UBound(m_columnWidths) Then
//!             cellValue = CStr(values(i))
//!             line = line & Left(cellValue & Space(m_columnWidths(i)), _
//!                                m_columnWidths(i))
//!         End If
//!     Next i
//!     
//!     ' Remove trailing spaces from line
//!     m_lines.Add RTrim(line)
//! End Sub
//!
//! Public Sub AddSeparator()
//!     Dim i As Integer
//!     Dim line As String
//!     
//!     For i = LBound(m_columnWidths) To UBound(m_columnWidths)
//!         line = line & String(m_columnWidths(i), "-")
//!     Next i
//!     
//!     m_lines.Add RTrim(line)
//! End Sub
//!
//! Public Function GetReport() As String
//!     Dim line As Variant
//!     Dim report As String
//!     
//!     For Each line In m_lines
//!         report = report & line & vbCrLf
//!     Next line
//!     
//!     GetReport = report
//! End Function
//!
//! Public Sub SaveToFile(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim line As Variant
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     For Each line In m_lines
//!         Print #fileNum, line
//!     Next line
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Example 4: Text Utilities Module
//! ```vb
//! ' Module: TextUtils
//!
//! Public Function CleanTrailing(ByVal text As String) As String
//!     CleanTrailing = RTrim(text)
//! End Function
//!
//! Public Function CleanBoth(ByVal text As String) As String
//!     CleanBoth = LTrim(RTrim(text))
//! End Function
//!
//! Public Function CleanArray(arr() As String, _
//!                           ByVal trimType As String) As String()
//!     Dim i As Integer
//!     Dim result() As String
//!     
//!     ReDim result(LBound(arr) To UBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         Select Case LCase(trimType)
//!             Case "trailing", "right"
//!                 result(i) = RTrim(arr(i))
//!             Case "leading", "left"
//!                 result(i) = LTrim(arr(i))
//!             Case "both", "all"
//!                 result(i) = LTrim(RTrim(arr(i)))
//!             Case Else
//!                 result(i) = arr(i)
//!         End Select
//!     Next i
//!     
//!     CleanArray = result
//! End Function
//!
//! Public Function PadAndTrim(ByVal text As String, _
//!                           ByVal width As Integer, _
//!                           Optional ByVal padChar As String = " ") As String
//!     Dim padded As String
//!     
//!     If Len(padChar) = 0 Then padChar = " "
//!     padded = Left(text & String(width, padChar), width)
//!     PadAndTrim = RTrim(padded)
//! End Function
//!
//! Public Function JoinClean(arr() As String, _
//!                          ByVal delimiter As String) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim cleanValue As String
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         cleanValue = RTrim(arr(i))
//!         If cleanValue <> "" Then
//!             If result <> "" Then
//!                 result = result & delimiter
//!             End If
//!             result = result & cleanValue
//!         End If
//!     Next i
//!     
//!     JoinClean = result
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' RTrim handles Null gracefully
//! Dim result As Variant
//! result = RTrim(Null)  ' Returns Null
//!
//! ' Safe trimming with Null check
//! Function SafeRTrim(ByVal value As Variant) As String
//!     If IsNull(value) Then
//!         SafeRTrim = ""
//!     Else
//!         SafeRTrim = RTrim(CStr(value))
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: String trimming is highly optimized
//! - **Creates New String**: Does not modify original (immutable)
//! - **Avoid in Tight Loops**: Cache result if using multiple times
//! - **Use `Trim()` for Both**: If removing both leading and trailing spaces
//!
//! ## Best Practices
//!
//! 1. **Use for database CHAR fields** - Remove padding from fixed-length fields
//! 2. **Clean export data** - Remove trailing spaces before writing to files
//! 3. **Combine with `LTrim`** - Use `Trim()` instead for both sides
//! 4. **Validate before use** - Check for Null if using Variant
//! 5. **Cache trimmed values** - Don't call repeatedly in loops
//! 6. **Apply to report output** - Clean formatting in generated reports
//! 7. **Use with fixed-width formats** - Essential for proper alignment
//! 8. **Document expectations** - Clarify if tabs/newlines should be removed
//! 9. **Test edge cases** - Empty strings, all spaces, Null values
//! 10. **Consider Unicode** - `RTrim` only removes ASCII space (32)
//!
//! ## Comparison with Related Functions
//!
//! | Function | Removes Leading | Removes Trailing | Removes Both |
//! |----------|----------------|------------------|--------------|
//! | **`RTrim`** | No | Yes | No |
//! | **`LTrim`** | Yes | No | No |
//! | **Trim** | Yes | Yes | Yes |
//!
//! ## `RTrim` vs `LTrim` vs Trim
//!
//! ```vb
//! Dim text As String
//! text = "   Hello World   "
//!
//! ' RTrim - removes trailing spaces only
//! Debug.Print "[" & RTrim(text) & "]"   ' [   Hello World]
//!
//! ' LTrim - removes leading spaces only
//! Debug.Print "[" & LTrim(text) & "]"   ' [Hello World   ]
//!
//! ' Trim - removes both leading and trailing
//! Debug.Print "[" & Trim(text) & "]"    ' [Hello World]
//!
//! ' Manual equivalent to Trim
//! Debug.Print "[" & LTrim(RTrim(text)) & "]"  ' [Hello World]
//! ```
//!
//! ## Whitespace Characters
//!
//! ```vb
//! ' RTrim only removes space (ASCII 32)
//! Dim text As String
//!
//! text = "Hello   "        ' Spaces - REMOVED
//! text = "Hello" & Chr(9)  ' Tab - NOT REMOVED
//! text = "Hello" & Chr(10) ' Line feed - NOT REMOVED
//! text = "Hello" & Chr(13) ' Carriage return - NOT REMOVED
//! text = "Hello" & Chr(160) ' Non-breaking space - NOT REMOVED
//!
//! ' To remove other whitespace, use custom function
//! Function TrimAllWhitespaceRight(ByVal text As String) As String
//!     Do While Len(text) > 0
//!         Dim ch As String
//!         ch = Right(text, 1)
//!         
//!         If ch = " " Or ch = Chr(9) Or ch = Chr(10) Or ch = Chr(13) Then
//!             text = Left(text, Len(text) - 1)
//!         Else
//!             Exit Do
//!         End If
//!     Loop
//!     
//!     TrimAllWhitespaceRight = text
//! End Function
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Works with ANSI and Unicode strings
//! - Only removes ASCII space character (32)
//! - Returns new string (original unchanged)
//! - Handles Null by returning Null
//! - Available in `VBScript`
//! - Same behavior across all Windows versions
//!
//! ## Limitations
//!
//! - **Only Space Character**: Does not remove tabs, line feeds, etc.
//! - **No Unicode Whitespace**: Does not remove non-breaking spaces, em spaces, etc.
//! - **Creates New String**: Cannot modify string in place
//! - **No Custom Characters**: Cannot specify which characters to remove
//! - **Null Propagation**: Returns Null if input is Null
//!
//! ## Related Functions
//!
//! - `LTrim`: Removes leading spaces from string
//! - `Trim`: Removes both leading and trailing spaces
//! - `Right`: Returns rightmost characters
//! - `Left`: Returns leftmost characters
//! - `Mid`: Returns substring from middle
//! - `Replace`: Replaces occurrences of substring
//! - `Space`: Creates string of spaces
//! - `Len`: Returns string length

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn rtrim_basic() {
        let source = r#"
            Dim result As String
            result = RTrim("Hello   ")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_variable() {
        let source = r#"
            cleaned = RTrim(userInput)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_database_field() {
        let source = r#"
            customerName = RTrim(rs("CustomerName"))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_if_statement() {
        let source = r#"
            If RTrim(text) = "" Then
                MsgBox "Empty"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_function_return() {
        let source = r#"
            Function CleanText(s As String) As String
                CleanText = RTrim(s)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_with_ltrim() {
        let source = r#"
            fullTrim = LTrim(RTrim(text))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_print_statement() {
        let source = r#"
            Print #1, RTrim(line)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_debug_print() {
        let source = r#"
            Debug.Print RTrim(text)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_with_statement() {
        let source = r#"
            With record
                .Name = RTrim(.Name)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_select_case() {
        let source = r#"
            Select Case RTrim(input)
                Case ""
                    MsgBox "Empty"
                Case Else
                    Process input
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_elseif() {
        let source = r#"
            If text = "" Then
                status = "Empty"
            ElseIf RTrim(text) = "" Then
                status = "Whitespace only"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_parentheses() {
        let source = r#"
            result = (RTrim(text))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_iif() {
        let source = r#"
            result = IIf(RTrim(text) = "", "Empty", "Has data")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_in_class() {
        let source = r#"
            Private Sub Class_Method()
                m_cleanValue = RTrim(m_rawValue)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_function_argument() {
        let source = r#"
            Call ProcessText(RTrim(input))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_property_assignment() {
        let source = r#"
            MyObject.CleanText = RTrim(dirtyText)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_array_assignment() {
        let source = r#"
            cleanValues(i) = RTrim(rawValues(i))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_for_loop() {
        let source = r#"
            For i = 1 To 10
                fields(i) = RTrim(fields(i))
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_while_wend() {
        let source = r#"
            While Not EOF(1)
                Line Input #1, line
                line = RTrim(line)
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_do_while() {
        let source = r#"
            Do While i < count
                text = RTrim(dataArray(i))
                i = i + 1
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_do_until() {
        let source = r#"
            Do Until RTrim(input) <> ""
                input = InputBox("Enter text")
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_msgbox() {
        let source = r#"
            MsgBox RTrim(message)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_concatenation() {
        let source = r#"
            reportLine = RTrim(customerName) & " - " & orderID
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_comparison() {
        let source = r#"
            If RTrim(text1) = RTrim(text2) Then
                MsgBox "Equal"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_label_caption() {
        let source = r#"
            lblName.Caption = RTrim(recordset("Name"))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_fixed_width() {
        let source = r#"
            outputLine = RTrim(Left(field & Space(20), 20))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rtrim_additem() {
        let source = r#"
            lstCustomers.AddItem RTrim(rs("CompanyName"))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RTrim"));
        assert!(text.contains("Identifier"));
    }
}
