//! # IsEmpty Function
//!
//! Returns a Boolean value indicating whether a Variant variable has been initialized.
//!
//! ## Syntax
//!
//! ```vb
//! IsEmpty(expression)
//! ```
//!
//! ## Parameters
//!
//! - `expression` (Required): Variant expression to test
//!
//! ## Return Value
//!
//! Returns a Boolean:
//! - `True` if the variable is Empty (uninitialized)
//! - `False` if the variable has been initialized
//! - Only Variant variables can be Empty
//! - All other variable types are always initialized with default values
//! - Returns `False` for Null values (Null is not the same as Empty)
//! - Returns `False` for zero, empty strings, and False
//!
//! ## Remarks
//!
//! The IsEmpty function is used to determine whether a Variant variable has been initialized:
//!
//! - Only works with Variant data type
//! - Returns `True` only for uninitialized Variant variables
//! - Empty is different from Null (IsNull) and zero
//! - Empty is different from zero-length string ("")
//! - Empty is different from False
//! - Useful for checking optional Variant parameters
//! - Can detect uninitialized elements in Variant arrays
//! - Once a Variant is assigned any value (including Null), it is no longer Empty
//! - Use VarType(var) = vbEmpty for the same check
//! - Cannot be used to test whether procedure or function exists
//! - Commonly used in procedures with optional Variant parameters
//!
//! ## Typical Uses
//!
//! 1. **Optional Parameters**: Check if optional Variant parameter was provided
//! 2. **Variable Initialization**: Verify Variant has been assigned a value
//! 3. **Array Elements**: Check if array elements have been assigned
//! 4. **Data Validation**: Distinguish between uninitialized and zero/empty values
//! 5. **Database Fields**: Detect uninitialized field values
//! 6. **Configuration Settings**: Check if settings have been loaded
//! 7. **Error Handling**: Detect uninitialized return values
//! 8. **Default Value Logic**: Apply defaults only when variable is uninitialized
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Check uninitialized Variant
//! Dim myVar As Variant
//!
//! If IsEmpty(myVar) Then
//!     Debug.Print "Variable is uninitialized"  ' This prints
//! End If
//!
//! myVar = 0
//! If IsEmpty(myVar) Then
//!     Debug.Print "Still empty"
//! Else
//!     Debug.Print "Now initialized"  ' This prints
//! End If
//!
//! ' Example 2: Distinguish Empty from other values
//! Dim testVar As Variant
//!
//! Debug.Print IsEmpty(testVar)        ' True - uninitialized
//! testVar = 0
//! Debug.Print IsEmpty(testVar)        ' False - initialized to zero
//! testVar = ""
//! Debug.Print IsEmpty(testVar)        ' False - initialized to empty string
//! testVar = False
//! Debug.Print IsEmpty(testVar)        ' False - initialized to False
//! testVar = Null
//! Debug.Print IsEmpty(testVar)        ' False - Null is not Empty
//!
//! ' Example 3: Optional parameter handling
//! Function ProcessData(data As String, Optional threshold As Variant) As Boolean
//!     Dim thresholdValue As Double
//!     
//!     If IsEmpty(threshold) Then
//!         ' Use default value when parameter not provided
//!         thresholdValue = 100
//!         Debug.Print "Using default threshold: 100"
//!     Else
//!         thresholdValue = threshold
//!         Debug.Print "Using provided threshold: " & threshold
//!     End If
//!     
//!     ' Process data with threshold
//!     ProcessData = (Len(data) > thresholdValue)
//! End Function
//!
//! ' Usage
//! ProcessData "test"              ' Uses default threshold (100)
//! ProcessData "test", 10          ' Uses provided threshold (10)
//! ProcessData "test", 0           ' Uses provided threshold (0) - not Empty!
//!
//! ' Example 4: Check array elements
//! Dim values(1 To 5) As Variant
//! Dim i As Integer
//!
//! values(1) = 10
//! values(3) = "Hello"
//! ' values(2), values(4), values(5) remain Empty
//!
//! For i = 1 To 5
//!     If IsEmpty(values(i)) Then
//!         Debug.Print "Element " & i & " is uninitialized"
//!     Else
//!         Debug.Print "Element " & i & " = " & values(i)
//!     End If
//! Next i
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Optional parameter with default
//! Function GetValue(key As String, Optional defaultValue As Variant) As Variant
//!     Dim result As Variant
//!     result = LookupValue(key)
//!     
//!     If IsEmpty(result) Then
//!         If IsEmpty(defaultValue) Then
//!             GetValue = Null  ' No default provided
//!         Else
//!             GetValue = defaultValue
//!         End If
//!     Else
//!         GetValue = result
//!     End If
//! End Function
//!
//! ' Pattern 2: Initialize if empty
//! Sub EnsureInitialized(ByRef value As Variant, defaultValue As Variant)
//!     If IsEmpty(value) Then
//!         value = defaultValue
//!     End If
//! End Sub
//!
//! ' Pattern 3: Count initialized array elements
//! Function CountInitializedElements(arr As Variant) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     If Not IsArray(arr) Then
//!         CountInitializedElements = 0
//!         Exit Function
//!     End If
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         If Not IsEmpty(arr(i)) Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountInitializedElements = count
//! End Function
//!
//! ' Pattern 4: Safe value retrieval
//! Function SafeGetValue(source As Variant, key As String) As Variant
//!     On Error Resume Next
//!     SafeGetValue = source(key)
//!     
//!     If Err.Number <> 0 Or IsEmpty(SafeGetValue) Then
//!         SafeGetValue = Null
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! ' Pattern 5: Validate required parameter
//! Sub ProcessRecord(recordData As Variant)
//!     If IsEmpty(recordData) Then
//!         Err.Raise 5, , "Record data is required"
//!     End If
//!     
//!     ' Process the record
//! End Sub
//!
//! ' Pattern 6: Type-aware value handling
//! Function DescribeValue(value As Variant) As String
//!     If IsEmpty(value) Then
//!         DescribeValue = "Uninitialized"
//!     ElseIf IsNull(value) Then
//!         DescribeValue = "Null value"
//!     ElseIf IsNumeric(value) Then
//!         DescribeValue = "Number: " & value
//!     ElseIf IsDate(value) Then
//!         DescribeValue = "Date: " & value
//!     Else
//!         DescribeValue = "Other: " & value
//!     End If
//! End Function
//!
//! ' Pattern 7: Clear variant (make it empty again)
//! Sub ClearVariant(ByRef value As Variant)
//!     ' Set to Empty state
//!     value = Empty
//!     Debug.Print "IsEmpty now: " & IsEmpty(value)  ' True
//! End Sub
//!
//! ' Pattern 8: Coalesce - return first non-empty value
//! Function Coalesce(ParamArray values() As Variant) As Variant
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If Not IsEmpty(values(i)) And Not IsNull(values(i)) Then
//!             Coalesce = values(i)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     Coalesce = Null  ' All values were Empty or Null
//! End Function
//!
//! ' Pattern 9: Check multiple values
//! Function AllInitialized(ParamArray values() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If IsEmpty(values(i)) Then
//!             AllInitialized = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     AllInitialized = True
//! End Function
//!
//! ' Pattern 10: Cache with lazy initialization
//! Private m_cachedValue As Variant
//!
//! Function GetCachedValue() As Variant
//!     If IsEmpty(m_cachedValue) Then
//!         ' Initialize cache on first access
//!         m_cachedValue = ExpensiveCalculation()
//!     End If
//!     
//!     GetCachedValue = m_cachedValue
//! End Function
//!
//! Sub InvalidateCache()
//!     m_cachedValue = Empty
//! End Sub
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Configuration manager with lazy loading
//! Public Class ConfigurationManager
//!     Private m_config As Variant
//!     Private m_filePath As String
//!     
//!     Public Sub Initialize(filePath As String)
//!         m_filePath = filePath
//!         m_config = Empty  ' Mark as uninitialized
//!     End Sub
//!     
//!     Public Function GetSetting(key As String, Optional defaultValue As Variant) As Variant
//!         ' Load config on first access
//!         If IsEmpty(m_config) Then
//!             LoadConfiguration
//!         End If
//!         
//!         On Error Resume Next
//!         GetSetting = m_config(key)
//!         
//!         If Err.Number <> 0 Or IsEmpty(GetSetting) Then
//!             If IsEmpty(defaultValue) Then
//!                 GetSetting = Null
//!             Else
//!                 GetSetting = defaultValue
//!             End If
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Private Sub LoadConfiguration()
//!         ' Load configuration from file
//!         Set m_config = CreateObject("Scripting.Dictionary")
//!         ' ... load settings into dictionary ...
//!     End Sub
//!     
//!     Public Sub Reload()
//!         m_config = Empty  ' Clear cache to force reload
//!     End Sub
//! End Class
//!
//! ' Example 2: Smart default value provider
//! Public Class DefaultValueProvider
//!     Private m_defaults As Variant
//!     
//!     Public Sub SetDefault(key As String, value As Variant)
//!         If IsEmpty(m_defaults) Then
//!             Set m_defaults = CreateObject("Scripting.Dictionary")
//!         End If
//!         m_defaults(key) = value
//!     End Sub
//!     
//!     Public Function GetWithDefault(value As Variant, key As String) As Variant
//!         If IsEmpty(value) Then
//!             If Not IsEmpty(m_defaults) Then
//!                 On Error Resume Next
//!                 GetWithDefault = m_defaults(key)
//!                 If Err.Number <> 0 Then
//!                     GetWithDefault = Null
//!                 End If
//!                 On Error GoTo 0
//!             Else
//!                 GetWithDefault = Null
//!             End If
//!         Else
//!             GetWithDefault = value
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 3: Variant array processor
//! Public Class VariantArrayProcessor
//!     Public Function Compact(arr As Variant) As Variant
//!         ' Remove Empty elements from array
//!         Dim result() As Variant
//!         Dim i As Long, count As Long
//!         
//!         If Not IsArray(arr) Then
//!             Compact = arr
//!             Exit Function
//!         End If
//!         
//!         ReDim result(LBound(arr) To UBound(arr))
//!         count = LBound(arr) - 1
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If Not IsEmpty(arr(i)) Then
//!                 count = count + 1
//!                 result(count) = arr(i)
//!             End If
//!         Next i
//!         
//!         If count >= LBound(result) Then
//!             ReDim Preserve result(LBound(result) To count)
//!             Compact = result
//!         Else
//!             Compact = Array()  ' Empty array
//!         End If
//!     End Function
//!     
//!     Public Function FillEmpty(arr As Variant, fillValue As Variant) As Variant
//!         ' Replace Empty elements with fill value
//!         Dim result() As Variant
//!         Dim i As Long
//!         
//!         If Not IsArray(arr) Then
//!             FillEmpty = arr
//!             Exit Function
//!         End If
//!         
//!         ReDim result(LBound(arr) To UBound(arr))
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If IsEmpty(arr(i)) Then
//!                 result(i) = fillValue
//!             Else
//!                 result(i) = arr(i)
//!             End If
//!         Next i
//!         
//!         FillEmpty = result
//!     End Function
//! End Class
//!
//! ' Example 4: Flexible function with multiple optional parameters
//! Function CreateReport(title As String, Optional author As Variant, _
//!                       Optional date As Variant, Optional includeIndex As Variant) As String
//!     Dim report As String
//!     
//!     report = "REPORT: " & title & vbCrLf
//!     
//!     If Not IsEmpty(author) Then
//!         report = report & "Author: " & author & vbCrLf
//!     Else
//!         report = report & "Author: Unknown" & vbCrLf
//!     End If
//!     
//!     If Not IsEmpty(date) Then
//!         If IsDate(date) Then
//!             report = report & "Date: " & Format$(date, "Long Date") & vbCrLf
//!         Else
//!             report = report & "Date: " & date & vbCrLf
//!         End If
//!     Else
//!         report = report & "Date: " & Format$(Now, "Long Date") & vbCrLf
//!     End If
//!     
//!     report = report & String(40, "-") & vbCrLf
//!     
//!     ' Add content...
//!     
//!     ' Add index if requested
//!     If Not IsEmpty(includeIndex) Then
//!         If includeIndex = True Then
//!             report = report & vbCrLf & "INDEX" & vbCrLf
//!             ' ... add index content ...
//!         End If
//!     End If
//!     
//!     CreateReport = report
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The IsEmpty function itself does not raise errors, but it's commonly used in error prevention:
//!
//! ```vb
//! Function SafeOperation(value As Variant) As Boolean
//!     If IsEmpty(value) Then
//!         MsgBox "Value must be initialized", vbExclamation
//!         SafeOperation = False
//!         Exit Function
//!     End If
//!     
//!     ' Safe to use value
//!     SafeOperation = True
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: IsEmpty is a very fast check with minimal overhead
//! - **Variant Only**: Only applicable to Variant variables
//! - **Early Validation**: Check IsEmpty early to avoid unnecessary processing
//! - **VarType Alternative**: `VarType(var) = vbEmpty` provides same check
//!
//! ## Best Practices
//!
//! 1. **Optional Parameters**: Always check IsEmpty for optional Variant parameters
//! 2. **Explicit Defaults**: Provide clear default values when parameters are Empty
//! 3. **Document Behavior**: Document whether Empty is valid for function parameters
//! 4. **Distinguish States**: Understand difference between Empty, Null, zero, and empty string
//! 5. **Initialization**: Consider explicitly initializing Variants when Empty is not desired
//! 6. **Combine Checks**: Use with IsNull for comprehensive validation
//! 7. **Clear Code**: Use IsEmpty rather than VarType for better readability
//! 8. **Reset to Empty**: Use `var = Empty` to reset Variant to uninitialized state
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | IsEmpty | Check if uninitialized | Boolean | Detect Empty Variants |
//! | IsNull | Check if Null | Boolean | Detect Null values |
//! | IsNumeric | Check if numeric | Boolean | Validate numeric data |
//! | IsDate | Check if date | Boolean | Validate date data |
//! | IsArray | Check if array | Boolean | Validate array variables |
//! | VarType | Get variant type | Integer | Detailed type information |
//! | TypeName | Get type name | String | Type name as string |
//!
//! ## Empty vs Null vs Zero vs Empty String
//!
//! ```vb
//! Dim v As Variant
//!
//! ' Empty (uninitialized)
//! Debug.Print IsEmpty(v)         ' True
//! Debug.Print IsNull(v)          ' False
//! Debug.Print v = 0              ' True (Empty coerces to 0 in numeric context)
//! Debug.Print v = ""             ' True (Empty coerces to "" in string context)
//!
//! ' Null
//! v = Null
//! Debug.Print IsEmpty(v)         ' False
//! Debug.Print IsNull(v)          ' True
//!
//! ' Zero
//! v = 0
//! Debug.Print IsEmpty(v)         ' False
//! Debug.Print v = 0              ' True
//!
//! ' Empty String
//! v = ""
//! Debug.Print IsEmpty(v)         ' False
//! Debug.Print v = ""             ' True
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns Boolean type
//! - Only works with Variant data type
//! - Empty is VB-specific concept (not in all languages)
//!
//! ## Limitations
//!
//! - Only works with Variant variables
//! - Cannot test non-Variant types (they're always initialized)
//! - Does not indicate what type of value is expected
//! - Cannot distinguish between intentionally Empty and accidentally uninitialized
//! - Empty has different coercion behavior in different contexts
//!
//! ## Related Functions
//!
//! - `IsNull`: Check if Variant is Null
//! - `VarType`: Get detailed Variant type information
//! - `TypeName`: Get type name as string
//! - `IsNumeric`: Check if numeric
//! - `IsDate`: Check if date
//! - `IsArray`: Check if array

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_isempty_basic() {
        let source = r#"
Sub Test()
    result = IsEmpty(myVariable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_if_statement() {
        let source = r#"
Sub Test()
    If IsEmpty(value) Then
        value = defaultValue
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_not_condition() {
        let source = r#"
Sub Test()
    If Not IsEmpty(param) Then
        ProcessValue param
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_function_return() {
        let source = r#"
Function CheckInitialized(v As Variant) As Boolean
    CheckInitialized = Not IsEmpty(v)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_boolean_and() {
        let source = r#"
Sub Test()
    If IsEmpty(value1) And IsEmpty(value2) Then
        InitializeDefaults
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_boolean_or() {
        let source = r#"
Sub Test()
    If IsEmpty(field) Or IsNull(field) Then
        ShowWarning
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_iif() {
        let source = r#"
Sub Test()
    displayValue = IIf(IsEmpty(value), "Not Set", value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is empty: " & IsEmpty(testVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Variable status: " & IsEmpty(myVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_do_while() {
        let source = r#"
Sub Test()
    Do While IsEmpty(cachedValue)
        cachedValue = LoadFromCache()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_do_until() {
        let source = r#"
Sub Test()
    Do Until Not IsEmpty(result)
        result = GetNextResult()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_variable_assignment() {
        let source = r#"
Sub Test()
    Dim isEmpty As Boolean
    isEmpty = IsEmpty(dataValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_property_assignment() {
        let source = r#"
Sub Test()
    obj.IsInitialized = Not IsEmpty(obj.Value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_isEmpty = IsEmpty(m_cachedData)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_with_statement() {
        let source = r#"
Sub Test()
    With config
        .RequiresInit = IsEmpty(.Settings)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_function_argument() {
        let source = r#"
Sub Test()
    Call ValidateParameter(IsEmpty(optionalParam))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_select_case() {
        let source = r#"
Sub Test()
    Select Case True
        Case IsEmpty(value)
            InitializeValue
        Case Else
            UseValue
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(arr)
        If IsEmpty(arr(i)) Then
            arr(i) = defaultValue
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_elseif() {
        let source = r#"
Sub Test()
    If IsNull(data) Then
        ProcessNull
    ElseIf IsEmpty(data) Then
        ProcessEmpty
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_concatenation() {
        let source = r#"
Sub Test()
    status = "Empty: " & IsEmpty(variable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_parentheses() {
        let source = r#"
Sub Test()
    result = (IsEmpty(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_array_check() {
        let source = r#"
Sub Test()
    checks(i) = IsEmpty(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_collection_add() {
        let source = r#"
Sub Test()
    states.Add IsEmpty(data(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_comparison() {
        let source = r#"
Sub Test()
    If IsEmpty(var1) = IsEmpty(var2) Then
        MsgBox "Same state"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IsEmpty(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_while_wend() {
        let source = r#"
Sub Test()
    While IsEmpty(buffer)
        buffer = ReadNext()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isempty_optional_param() {
        let source = r#"
Function Process(Optional param As Variant) As Boolean
    If IsEmpty(param) Then
        param = GetDefaultValue()
    End If
    Process = True
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsEmpty"));
        assert!(text.contains("Identifier"));
    }
}
