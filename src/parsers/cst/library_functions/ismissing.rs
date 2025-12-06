//! # `IsMissing` Function
//!
//! Returns a `Boolean` value indicating whether an optional `Variant` parameter was passed to a procedure.
//!
//! ## Syntax
//!
//! ```vb
//! IsMissing(argname)
//! ```
//!
//! ## Parameters
//!
//! - `argname` (Required): Name of an optional `Variant` parameter
//!
//! ## Return Value
//!
//! Returns a Boolean:
//! - `True` if the optional `Variant` argument was not passed
//! - `False` if the optional `Variant` argument was passed
//! - Only works with optional `Variant` parameters
//! - Does not work with other data types (only `Variant`)
//! - Does not work with `ParamArray` parameters
//! - Returns `False` if `Null`, `Empty`, or any value was explicitly passed
//!
//! ## Remarks
//!
//! The `IsMissing` function is used to detect whether an optional parameter was omitted:
//!
//! - Only works with Optional `Variant` parameters
//! - Cannot be used with typed optional parameters (Integer, String, etc.)
//! - Returns `True` only when argument was completely omitted
//! - Returns `False` if any value was passed (including `Empty`, `Null`, 0, "")
//! - Useful for implementing functions with truly optional behavior
//! - Different from checking `IsEmpty` - `IsMissing` detects omission
//! - `Empty` can be explicitly passed: `MyFunc Empty` - `IsMissing` returns `False`
//! - Common in COM/ActiveX programming for optional parameters
//! - Allows distinguishing "not provided" from "provided as Empty/Null/0"
//! - Must be used directly on parameter name, not on expressions
//! - Parameter must be declared as Optional `Variant`
//! - Cannot be used after assigning the parameter to another variable
//!
//! ## Typical Uses
//!
//! 1. **Optional Parameter Detection**: Check if optional argument was provided
//! 2. **Default Value Logic**: Apply different defaults based on whether parameter was omitted
//! 3. **API Compatibility**: Maintain backward compatibility with varying parameter counts
//! 4. **COM Interop**: Work with COM objects expecting optional parameters
//! 5. **Flexible Functions**: Create functions with multiple optional behaviors
//! 6. **Database Operations**: Handle optional `WHERE` clause parameters
//! 7. **Configuration Functions**: Apply settings only when explicitly provided
//! 8. **Validation Logic**: Distinguish "no value" from "zero value"
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple optional parameter handling
//! Function Greet(name As String, Optional title As Variant) As String
//!     Dim greeting As String
//!     
//!     If IsMissing(title) Then
//!         greeting = "Hello, " & name
//!     Else
//!         greeting = "Hello, " & title & " " & name
//!     End If
//!     
//!     Greet = greeting
//! End Function
//!
//! ' Usage
//! Debug.Print Greet("Smith")              ' "Hello, Smith"
//! Debug.Print Greet("Smith", "Dr.")       ' "Hello, Dr. Smith"
//! Debug.Print Greet("Smith", Empty)       ' "Hello,  Smith" - Empty was passed!
//!
//! ' Example 2: Multiple optional parameters
//! Sub LogMessage(msg As String, Optional level As Variant, Optional timestamp As Variant)
//!     Dim output As String
//!     
//!     output = msg
//!     
//!     If Not IsMissing(level) Then
//!         output = "[" & level & "] " & output
//!     End If
//!     
//!     If Not IsMissing(timestamp) Then
//!         output = Format$(timestamp, "hh:nn:ss") & " " & output
//!     Else
//!         output = Format$(Now, "hh:nn:ss") & " " & output
//!     End If
//!     
//!     Debug.Print output
//! End Sub
//!
//! ' Usage
//! LogMessage "Application started"                    ' Uses current time, no level
//! LogMessage "Error occurred", "ERROR"                ' Uses current time, ERROR level
//! LogMessage "Debug info", "DEBUG", #1/1/2025 10:30#  ' Uses specified time
//!
//! ' Example 3: Distinguish missing from zero/empty
//! Function Calculate(value As Double, Optional multiplier As Variant) As Double
//!     If IsMissing(multiplier) Then
//!         ' No multiplier provided - return original value
//!         Calculate = value
//!     ElseIf multiplier = 0 Then
//!         ' Multiplier is explicitly 0 - return 0
//!         Calculate = 0
//!     Else
//!         ' Multiplier provided and non-zero
//!         Calculate = value * multiplier
//!     End If
//! End Function
//!
//! Debug.Print Calculate(10)           ' 10 - multiplier missing
//! Debug.Print Calculate(10, 2)        ' 20 - multiplier is 2
//! Debug.Print Calculate(10, 0)        ' 0 - multiplier is explicitly 0
//!
//! ' Example 4: Database query with optional filter
//! Function GetRecords(table As String, Optional whereClause As Variant) As Recordset
//!     Dim sql As String
//!     
//!     sql = "SELECT * FROM " & table
//!     
//!     If Not IsMissing(whereClause) Then
//!         If whereClause <> "" Then
//!             sql = sql & " WHERE " & whereClause
//!         End If
//!     End If
//!     
//!     Set GetRecords = db.OpenRecordset(sql)
//! End Function
//!
//! ' Usage
//! Set rs = GetRecords("Customers")                    ' All records
//! Set rs = GetRecords("Customers", "State = 'CA'")    ' Filtered records
//! Set rs = GetRecords("Customers", "")                ' All records (empty string passed)
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Optional parameter with computed default
//! Function ProcessData(data As String, Optional maxLength As Variant) As String
//!     Dim limit As Long
//!     
//!     If IsMissing(maxLength) Then
//!         limit = Len(data)  ' Use full length when not specified
//!     Else
//!         limit = maxLength
//!     End If
//!     
//!     ProcessData = Left$(data, limit)
//! End Function
//!
//! ' Pattern 2: Cascading optional parameters
//! Sub SaveFile(filename As String, Optional path As Variant, Optional createBackup As Variant)
//!     Dim fullPath As String
//!     Dim backup As Boolean
//!     
//!     If IsMissing(path) Then
//!         fullPath = App.Path & "\" & filename
//!     Else
//!         fullPath = path & "\" & filename
//!     End If
//!     
//!     If IsMissing(createBackup) Then
//!         backup = False  ' Default to no backup
//!     Else
//!         backup = createBackup
//!     End If
//!     
//!     ' Save file logic...
//! End Sub
//!
//! ' Pattern 3: Count provided optional parameters
//! Function CountProvided(Optional arg1 As Variant, Optional arg2 As Variant, _
//!                        Optional arg3 As Variant) As Integer
//!     Dim count As Integer
//!     
//!     count = 0
//!     If Not IsMissing(arg1) Then count = count + 1
//!     If Not IsMissing(arg2) Then count = count + 1
//!     If Not IsMissing(arg3) Then count = count + 1
//!     
//!     CountProvided = count
//! End Function
//!
//! ' Pattern 4: Optional override parameter
//! Function GetSetting(key As String, Optional overrideValue As Variant) As Variant
//!     If IsMissing(overrideValue) Then
//!         ' Load from registry or config file
//!         GetSetting = LoadFromConfig(key)
//!     Else
//!         ' Use provided override
//!         GetSetting = overrideValue
//!     End If
//! End Function
//!
//! ' Pattern 5: Optional parameter affects behavior
//! Sub PrintReport(data As Variant, Optional includeHeader As Variant)
//!     If Not IsMissing(includeHeader) Then
//!         If includeHeader Then
//!             PrintHeader
//!         End If
//!     Else
//!         ' Default: always include header when not specified
//!         PrintHeader
//!     End If
//!     
//!     PrintData data
//! End Sub
//!
//! ' Pattern 6: Validation with optional strict mode
//! Function ValidateEmail(email As String, Optional strictMode As Variant) As Boolean
//!     Dim strict As Boolean
//!     
//!     If IsMissing(strictMode) Then
//!         strict = False  ' Default to lenient validation
//!     Else
//!         strict = strictMode
//!     End If
//!     
//!     If strict Then
//!         ValidateEmail = ValidateEmailStrict(email)
//!     Else
//!         ValidateEmail = ValidateEmailBasic(email)
//!     End If
//! End Function
//!
//! ' Pattern 7: Optional range parameters
//! Function GetSubstring(text As String, Optional startPos As Variant, _
//!                       Optional length As Variant) As String
//!     Dim start As Long
//!     Dim len As Long
//!     
//!     If IsMissing(startPos) Then
//!         start = 1
//!     Else
//!         start = startPos
//!     End If
//!     
//!     If IsMissing(length) Then
//!         len = Len(text) - start + 1
//!     Else
//!         len = length
//!     End If
//!     
//!     GetSubstring = Mid$(text, start, len)
//! End Function
//!
//! ' Pattern 8: Build parameter list dynamically
//! Function BuildCommand(command As String, Optional arg1 As Variant, _
//!                       Optional arg2 As Variant) As String
//!     Dim cmd As String
//!     
//!     cmd = command
//!     
//!     If Not IsMissing(arg1) Then
//!         cmd = cmd & " " & arg1
//!     End If
//!     
//!     If Not IsMissing(arg2) Then
//!         cmd = cmd & " " & arg2
//!     End If
//!     
//!     BuildCommand = cmd
//! End Function
//!
//! ' Pattern 9: Optional error handler callback
//! Function ProcessRecords(records As Variant, Optional errorHandler As Variant) As Long
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     count = 0
//!     For i = LBound(records) To UBound(records)
//!         On Error Resume Next
//!         ProcessRecord records(i)
//!         
//!         If Err.Number <> 0 Then
//!             If Not IsMissing(errorHandler) Then
//!                 ' Call custom error handler if provided
//!                 Application.Run errorHandler, records(i), Err.Number
//!             End If
//!             Err.Clear
//!         Else
//!             count = count + 1
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     ProcessRecords = count
//! End Function
//!
//! ' Pattern 10: Optional configuration object
//! Sub Initialize(Optional config As Variant)
//!     If IsMissing(config) Then
//!         ' Use default configuration
//!         LoadDefaultConfig
//!     Else
//!         ' Apply provided configuration
//!         ApplyConfig config
//!     End If
//! End Sub
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Flexible query builder
//! Public Class QueryBuilder
//!     Private m_sql As String
//!     
//!     Public Sub SelectFrom(table As String, Optional fields As Variant, _
//!                          Optional whereClause As Variant, _
//!                          Optional orderBy As Variant)
//!         If IsMissing(fields) Then
//!             m_sql = "SELECT * FROM " & table
//!         Else
//!             m_sql = "SELECT " & fields & " FROM " & table
//!         End If
//!         
//!         If Not IsMissing(whereClause) Then
//!             If whereClause <> "" Then
//!                 m_sql = m_sql & " WHERE " & whereClause
//!             End If
//!         End If
//!         
//!         If Not IsMissing(orderBy) Then
//!             If orderBy <> "" Then
//!                 m_sql = m_sql & " ORDER BY " & orderBy
//!             End If
//!         End If
//!     End Sub
//!     
//!     Public Function GetSQL() As String
//!         GetSQL = m_sql
//!     End Function
//! End Class
//!
//! ' Usage:
//! Dim qb As New QueryBuilder
//! qb.SelectFrom "Customers"                           ' SELECT * FROM Customers
//! qb.SelectFrom "Customers", "Name, Email"            ' SELECT Name, Email FROM Customers
//! qb.SelectFrom "Customers", , "State = 'CA'"         ' SELECT * FROM Customers WHERE State = 'CA'
//! qb.SelectFrom "Customers", "Name", "Active = 1", "Name"  ' Full query
//!
//! ' Example 2: Logger with flexible output
//! Public Class Logger
//!     Private m_logFile As String
//!     
//!     Public Sub Initialize(Optional filename As Variant)
//!         If IsMissing(filename) Then
//!             m_logFile = App.Path & "\app.log"
//!         Else
//!             m_logFile = filename
//!         End If
//!     End Sub
//!     
//!     Public Sub Log(message As String, Optional level As Variant, _
//!                    Optional timestamp As Variant, Optional writeToFile As Variant)
//!         Dim output As String
//!         Dim logLevel As String
//!         Dim useTimestamp As Date
//!         Dim toFile As Boolean
//!         
//!         ' Determine log level
//!         If IsMissing(level) Then
//!             logLevel = "INFO"
//!         Else
//!             logLevel = level
//!         End If
//!         
//!         ' Determine timestamp
//!         If IsMissing(timestamp) Then
//!             useTimestamp = Now
//!         Else
//!             useTimestamp = timestamp
//!         End If
//!         
//!         ' Determine output destination
//!         If IsMissing(writeToFile) Then
//!             toFile = True  ' Default to file
//!         Else
//!             toFile = writeToFile
//!         End If
//!         
//!         ' Build output
//!         output = Format$(useTimestamp, "yyyy-mm-dd hh:nn:ss") & " [" & _
//!                  logLevel & "] " & message
//!         
//!         ' Write output
//!         Debug.Print output
//!         If toFile Then
//!             WriteToLogFile output
//!         End If
//!     End Sub
//!     
//!     Private Sub WriteToLogFile(text As String)
//!         Dim fileNum As Integer
//!         fileNum = FreeFile
//!         Open m_logFile For Append As fileNum
//!         Print #fileNum, text
//!         Close fileNum
//!     End Sub
//! End Class
//!
//! ' Example 3: HTTP request builder
//! Public Class HttpRequest
//!     Public Function Get(url As String, Optional headers As Variant, _
//!                        Optional timeout As Variant) As String
//!         Dim http As Object
//!         Set http = CreateObject("MSXML2.XMLHTTP")
//!         
//!         ' Set timeout if provided
//!         If Not IsMissing(timeout) Then
//!             http.SetTimeouts timeout, timeout, timeout, timeout
//!         End If
//!         
//!         http.Open "GET", url, False
//!         
//!         ' Add custom headers if provided
//!         If Not IsMissing(headers) Then
//!             Dim headerList As Variant
//!             Dim i As Long
//!             
//!             If IsArray(headers) Then
//!                 For i = LBound(headers) To UBound(headers)
//!                     http.setRequestHeader Split(headers(i), ":")(0), _
//!                                          Trim$(Split(headers(i), ":")(1))
//!                 Next i
//!             End If
//!         End If
//!         
//!         http.send
//!         Get = http.responseText
//!     End Function
//! End Class
//!
//! ' Example 4: Validation framework
//! Public Class Validator
//!     Public Function Validate(value As Variant, Optional minValue As Variant, _
//!                             Optional maxValue As Variant, _
//!                             Optional pattern As Variant) As Boolean
//!         Dim isValid As Boolean
//!         isValid = True
//!         
//!         ' Check minimum value
//!         If Not IsMissing(minValue) Then
//!             If value < minValue Then
//!                 isValid = False
//!                 Exit Function
//!             End If
//!         End If
//!         
//!         ' Check maximum value
//!         If Not IsMissing(maxValue) Then
//!             If value > maxValue Then
//!                 isValid = False
//!                 Exit Function
//!             End If
//!         End If
//!         
//!         ' Check pattern (if string)
//!         If Not IsMissing(pattern) Then
//!             If VarType(value) = vbString Then
//!                 isValid = MatchesPattern(CStr(value), CStr(pattern))
//!             End If
//!         End If
//!         
//!         Validate = isValid
//!     End Function
//!     
//!     Private Function MatchesPattern(text As String, pattern As String) As Boolean
//!         ' Simple pattern matching implementation
//!         MatchesPattern = (text Like pattern)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! `IsMissing` itself does not raise errors, but improper usage can:
//!
//! ```vb
//! ' ERROR: Cannot use with non-Variant optional parameters
//! Function BadExample(Optional value As Integer) As Boolean
//!     BadExample = IsMissing(value)  ' Compile error!
//! End Function
//!
//! ' CORRECT: Must use Optional Variant
//! Function GoodExample(Optional value As Variant) As Boolean
//!     GoodExample = IsMissing(value)  ' Works correctly
//! End Function
//!
//! ' ERROR: Cannot use after assignment
//! Function BadExample2(Optional value As Variant) As Boolean
//!     Dim temp As Variant
//!     temp = value
//!     BadExample2 = IsMissing(temp)  ' Always False! Use value directly.
//! End Function
//!
//! ' CORRECT: Use parameter directly
//! Function GoodExample2(Optional value As Variant) As Boolean
//!     GoodExample2 = IsMissing(value)  ' Works correctly
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsMissing` is a very fast check with minimal overhead
//! - **Compile-Time Check**: VB6 can optimize `IsMissing` checks
//! - **Variant Overhead**: Optional `Variant` parameters have more overhead than typed parameters
//! - **Use Sparingly**: Only use when you truly need to distinguish missing from provided
//!
//! ## Best Practices
//!
//! 1. **Use Only with Variant**: `IsMissing` only works with Optional `Variant` parameters
//! 2. **Direct Check**: Always check `IsMissing` directly on the parameter name
//! 3. **Document Behavior**: Clearly document what happens when parameter is omitted
//! 4. **Provide Defaults**: Consider if default values on Optional parameters would work instead
//! 5. **Avoid Complexity**: Don't overuse optional parameters - can make APIs confusing
//! 6. **Check Early**: Test `IsMissing` before using the parameter value
//! 7. **Combine Wisely**: Can combine with `IsEmpty`, `IsNull` checks for complete validation
//! 8. **API Design**: Use for true optional behavior, not just to avoid typing
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsMissing` | Check if optional parameter omitted | `Boolean` | Detect missing Optional `Variant` arguments |
//! | `IsEmpty` | Check if uninitialized | `Boolean` | Detect `Empty` `Variant` values |
//! | `IsNull` | Check if Null | `Boolean` | Detect `Null` values |
//! | `IsError` | Check if error value | `Boolean` | Detect `CVErr` error values |
//! | `VarType` | Get variant type | `Integer` | Detailed type information |
//! | `TypeName` | Get type name | `String` | Type name as string |
//!
//! ## `IsMissing` vs `IsEmpty`
//!
//! ```vb
//! Sub Test(Optional param As Variant)
//!     ' Case 1: Parameter not provided
//!     ' Call: Test
//!     Debug.Print IsMissing(param)  ' True - was not provided
//!     Debug.Print IsEmpty(param)    ' True - is Empty
//!     
//!     ' Case 2: Explicit Empty passed
//!     ' Call: Test Empty
//!     Debug.Print IsMissing(param)  ' False - was provided (even though Empty)
//!     Debug.Print IsEmpty(param)    ' True - is Empty
//!     
//!     ' Case 3: Null passed
//!     ' Call: Test Null
//!     Debug.Print IsMissing(param)  ' False - was provided
//!     Debug.Print IsEmpty(param)    ' False - is Null, not Empty
//!     
//!     ' Case 4: Zero passed
//!     ' Call: Test 0
//!     Debug.Print IsMissing(param)  ' False - was provided
//!     Debug.Print IsEmpty(param)    ' False - has value (0)
//! End Sub
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Only works with Optional `Variant` parameters
//! - Compile-time requirement: parameter must be Optional `Variant`
//! - Cannot be used with `ParamArray`
//! - Common in COM/ActiveX programming
//!
//! ## Limitations
//!
//! - Only works with Optional `Variant` parameters (not `Integer`, `String`, etc.)
//! - Cannot be used on expressions, only on parameter names
//! - Cannot be used after assigning parameter to another variable
//! - Does not work with `ParamArray` parameters
//! - Cannot detect which parameter in `ParamArray` was omitted
//! - Requires `Variant` type (introduces type safety concerns)
//! - Can make function signatures more complex
//!
//! ## Related Functions
//!
//! - `IsEmpty`: Check if `Variant` is uninitialized (`Empty`)
//! - `IsNull`: Check if `Variant` is `Null`
//! - `VarType`: Get detailed `Variant` type information
//! - `TypeName`: Get type name as `String`

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ismissing_basic() {
        let source = r#"
Sub Test(Optional param As Variant)
    result = IsMissing(param)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_if_statement() {
        let source = r#"
Function Greet(name As String, Optional title As Variant) As String
    If IsMissing(title) Then
        Greet = "Hello " & name
    Else
        Greet = "Hello " & title & " " & name
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_not_condition() {
        let source = r#"
Sub Test(Optional value As Variant)
    If Not IsMissing(value) Then
        ProcessValue value
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_function_return() {
        let source = r#"
Function IsProvided(Optional arg As Variant) As Boolean
    IsProvided = Not IsMissing(arg)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_boolean_and() {
        let source = r#"
Sub Test(Optional arg1 As Variant, Optional arg2 As Variant)
    If IsMissing(arg1) And IsMissing(arg2) Then
        UseDefaults
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_boolean_or() {
        let source = r#"
Sub Test(Optional param1 As Variant, Optional param2 As Variant)
    If IsMissing(param1) Or IsMissing(param2) Then
        ShowWarning
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_iif() {
        let source = r#"
Function GetValue(Optional value As Variant) As String
    GetValue = IIf(IsMissing(value), "Default", value)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_debug_print() {
        let source = r#"
Sub Test(Optional arg As Variant)
    Debug.Print "Missing: " & IsMissing(arg)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_msgbox() {
        let source = r#"
Sub Test(Optional myParam As Variant)
    MsgBox "Parameter status: " & IsMissing(myParam)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_do_while() {
        let source = r#"
Sub Test(Optional config As Variant)
    Do While IsMissing(config)
        config = GetDefaultConfig()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_do_until() {
        let source = r#"
Sub Test(Optional setting As Variant)
    Do Until Not IsMissing(setting)
        setting = PromptForSetting()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_variable_assignment() {
        let source = r#"
Sub Test(Optional data As Variant)
    Dim isMissing As Boolean
    isMissing = IsMissing(data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_property_assignment() {
        let source = r#"
Sub Test(Optional opt As Variant)
    obj.WasProvided = Not IsMissing(opt)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_in_class() {
        let source = r#"
Public Sub Initialize(Optional settings As Variant)
    m_hasSettings = Not IsMissing(settings)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_with_statement() {
        let source = r#"
Sub Configure(Optional options As Variant)
    With config
        .UseDefaults = IsMissing(options)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_function_argument() {
        let source = r#"
Sub Test(Optional param As Variant)
    Call LogStatus(IsMissing(param))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_select_case() {
        let source = r#"
Sub Process(Optional mode As Variant)
    Select Case True
        Case IsMissing(mode)
            UseDefaultMode
        Case Else
            UseSpecifiedMode mode
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_for_loop() {
        let source = r#"
Function CountProvided(Optional a As Variant, Optional b As Variant, Optional c As Variant) As Integer
    Dim count As Integer
    Dim params(0 To 2) As Variant
    Dim i As Integer
    
    If Not IsMissing(a) Then count = count + 1
    CountProvided = count
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_elseif() {
        let source = r#"
Sub Handle(Optional param As Variant)
    If IsNull(param) Then
        ProcessNull
    ElseIf IsMissing(param) Then
        ProcessMissing
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_concatenation() {
        let source = r#"
Sub Report(Optional value As Variant)
    status = "Provided: " & Not IsMissing(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_parentheses() {
        let source = r#"
Sub Test(Optional arg As Variant)
    result = (IsMissing(arg))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_multiple_params() {
        let source = r#"
Sub SaveFile(filename As String, Optional path As Variant, Optional backup As Variant)
    If IsMissing(path) Then
        path = App.Path
    End If
    If IsMissing(backup) Then
        backup = False
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_collection_check() {
        let source = r#"
Sub AddOptional(coll As Collection, Optional item As Variant)
    If Not IsMissing(item) Then
        coll.Add item
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_comparison() {
        let source = r#"
Function Compare(Optional a As Variant, Optional b As Variant) As Boolean
    Compare = (IsMissing(a) = IsMissing(b))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_nested_call() {
        let source = r#"
Sub Test(Optional value As Variant)
    result = CStr(IsMissing(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_while_wend() {
        let source = r#"
Sub Test(Optional input As Variant)
    While IsMissing(input)
        input = GetUserInput()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ismissing_default_logic() {
        let source = r#"
Function Calculate(x As Double, Optional multiplier As Variant) As Double
    If IsMissing(multiplier) Then
        Calculate = x
    Else
        Calculate = x * multiplier
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsMissing"));
        assert!(text.contains("Identifier"));
    }
}
