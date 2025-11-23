//! # IsObject Function
//!
//! Returns a Boolean value indicating whether an identifier represents an object variable.
//!
//! ## Syntax
//!
//! ```vb
//! IsObject(identifier)
//! ```
//!
//! ## Parameters
//!
//! - `identifier` (Required): Variable name to test
//!
//! ## Return Value
//!
//! Returns a Boolean:
//! - `True` if the identifier is an object variable (Object, Form, Control, etc.)
//! - `False` if the identifier is not an object variable
//! - Returns `True` for any object reference (including Nothing)
//! - Returns `False` for numeric types, strings, dates, arrays of non-objects
//! - Returns `False` for Null and Empty
//! - Works with Variant variables containing object references
//! - Use to determine if variable can be used with Set statement
//!
//! ## Remarks
//!
//! The IsObject function is used to determine whether a variable is an object reference:
//!
//! - Returns True for any object variable type (Form, Control, Collection, etc.)
//! - Returns True even if object is Nothing (uninitialized object reference)
//! - Returns False for value types (Integer, String, Double, Boolean, etc.)
//! - Returns False for Null and Empty
//! - Useful before calling object methods or properties
//! - Use to determine if Set statement is needed for assignment
//! - Common in COM/ActiveX programming
//! - Works with Variant variables containing object references
//! - Cannot distinguish between different object types (use TypeOf...Is for that)
//! - VarType(var) = vbObject provides similar but not identical information
//! - Important for proper cleanup (setting objects to Nothing)
//! - Use before accessing object members to avoid errors
//!
//! ## Typical Uses
//!
//! 1. **Object Detection**: Check if variable contains an object reference
//! 2. **Assignment Logic**: Determine whether to use Set or regular assignment
//! 3. **Cleanup Verification**: Verify objects are properly set to Nothing
//! 4. **Error Prevention**: Avoid "Object required" errors
//! 5. **Type Validation**: Validate function parameters are objects
//! 6. **COM Interop**: Work with late-bound COM objects
//! 7. **Collection Processing**: Handle mixed collections of objects and values
//! 8. **Variant Handling**: Determine if Variant contains object or value
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic object detection
//! Dim obj As Object
//! Dim str As String
//! Dim num As Integer
//!
//! Set obj = CreateObject("Scripting.Dictionary")
//! str = "Hello"
//! num = 123
//!
//! Debug.Print IsObject(obj)      ' True - object variable
//! Debug.Print IsObject(str)      ' False - string variable
//! Debug.Print IsObject(num)      ' False - numeric variable
//!
//! ' Example 2: Nothing is still an object type
//! Dim myObj As Object
//!
//! Set myObj = Nothing
//! Debug.Print IsObject(myObj)    ' True - object type (even though Nothing)
//!
//! ' Example 3: Various object types
//! Dim frm As Form
//! Dim col As Collection
//! Dim dict As Object
//!
//! Set frm = New Form1
//! Set col = New Collection
//! Set dict = CreateObject("Scripting.Dictionary")
//!
//! Debug.Print IsObject(frm)      ' True - Form object
//! Debug.Print IsObject(col)      ' True - Collection object
//! Debug.Print IsObject(dict)     ' True - Dictionary object
//!
//! ' Example 4: Determine assignment method
//! Sub AssignValue(target As Variant, source As Variant)
//!     If IsObject(source) Then
//!         Set target = source    ' Use Set for objects
//!     Else
//!         target = source        ' Regular assignment for values
//!     End If
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe object cleanup
//! Sub CleanupObject(ByRef obj As Variant)
//!     If IsObject(obj) Then
//!         Set obj = Nothing
//!     End If
//! End Sub
//!
//! ' Pattern 2: Check if object is set before use
//! Function SafeCallMethod(obj As Variant) As Boolean
//!     If IsObject(obj) Then
//!         If Not obj Is Nothing Then
//!             obj.SomeMethod
//!             SafeCallMethod = True
//!         Else
//!             SafeCallMethod = False
//!         End If
//!     Else
//!         SafeCallMethod = False
//!     End If
//! End Function
//!
//! ' Pattern 3: Copy with correct assignment
//! Function CopyValue(source As Variant) As Variant
//!     If IsObject(source) Then
//!         Set CopyValue = source
//!     Else
//!         CopyValue = source
//!     End If
//! End Function
//!
//! ' Pattern 4: Count objects in collection
//! Function CountObjects(items As Collection) As Long
//!     Dim count As Long
//!     Dim item As Variant
//!     
//!     count = 0
//!     For Each item In items
//!         If IsObject(item) Then
//!             count = count + 1
//!         End If
//!     Next item
//!     
//!     CountObjects = count
//! End Function
//!
//! ' Pattern 5: Validate function parameter
//! Function ProcessObject(obj As Variant) As Boolean
//!     If Not IsObject(obj) Then
//!         Err.Raise 5, , "Object parameter required"
//!     End If
//!     
//!     If obj Is Nothing Then
//!         Err.Raise 91, , "Object not set"
//!     End If
//!     
//!     ' Process object
//!     ProcessObject = True
//! End Function
//!
//! ' Pattern 6: Generic value display
//! Function ValueToString(value As Variant) As String
//!     If IsNull(value) Then
//!         ValueToString = "Null"
//!     ElseIf IsEmpty(value) Then
//!         ValueToString = "Empty"
//!     ElseIf IsObject(value) Then
//!         If value Is Nothing Then
//!             ValueToString = "Nothing"
//!         Else
//!             ValueToString = TypeName(value) & " object"
//!         End If
//!     Else
//!         ValueToString = CStr(value)
//!     End If
//! End Function
//!
//! ' Pattern 7: Clone array with proper assignment
//! Function CloneArray(arr As Variant) As Variant
//!     Dim result() As Variant
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then
//!         CloneArray = arr
//!         Exit Function
//!     End If
//!     
//!     ReDim result(LBound(arr) To UBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If IsObject(arr(i)) Then
//!             Set result(i) = arr(i)
//!         Else
//!             result(i) = arr(i)
//!         End If
//!     Next i
//!     
//!     CloneArray = result
//! End Function
//!
//! ' Pattern 8: Release all objects in array
//! Sub ReleaseObjects(arr As Variant)
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then Exit Sub
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If IsObject(arr(i)) Then
//!             Set arr(i) = Nothing
//!         End If
//!     Next i
//! End Sub
//!
//! ' Pattern 9: Filter objects from collection
//! Function GetObjects(items As Collection) As Collection
//!     Dim result As New Collection
//!     Dim item As Variant
//!     
//!     For Each item In items
//!         If IsObject(item) Then
//!             result.Add item
//!         End If
//!     Next item
//!     
//!     Set GetObjects = result
//! End Function
//!
//! ' Pattern 10: Compare values with object handling
//! Function AreEqual(val1 As Variant, val2 As Variant) As Boolean
//!     If IsObject(val1) And IsObject(val2) Then
//!         AreEqual = (val1 Is val2)  ' Use Is for object comparison
//!     ElseIf IsObject(val1) Or IsObject(val2) Then
//!         AreEqual = False  ' One object, one value - not equal
//!     Else
//!         AreEqual = (val1 = val2)  ' Regular comparison for values
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Generic collection wrapper
//! Public Class GenericCollection
//!     Private m_items As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_items = New Collection
//!     End Sub
//!     
//!     Public Sub Add(item As Variant)
//!         If IsObject(item) Then
//!             m_items.Add item
//!         Else
//!             m_items.Add item
//!         End If
//!     End Sub
//!     
//!     Public Function Item(index As Variant) As Variant
//!         If IsObject(m_items(index)) Then
//!             Set Item = m_items(index)
//!         Else
//!             Item = m_items(index)
//!         End If
//!     End Function
//!     
//!     Public Sub Clear()
//!         Dim i As Long
//!         
//!         For i = m_items.Count To 1 Step -1
//!             If IsObject(m_items(i)) Then
//!                 Set m_items(i) = Nothing
//!             End If
//!         Next i
//!         
//!         Set m_items = New Collection
//!     End Sub
//!     
//!     Public Function Count() As Long
//!         Count = m_items.Count
//!     End Function
//! End Class
//!
//! ' Example 2: Object lifecycle manager
//! Public Class ObjectManager
//!     Private m_objects As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_objects = New Collection
//!     End Sub
//!     
//!     Public Sub Register(obj As Variant, Optional key As String = "")
//!         If Not IsObject(obj) Then
//!             Err.Raise 5, , "Only objects can be registered"
//!         End If
//!         
//!         If key = "" Then
//!             m_objects.Add obj
//!         Else
//!             m_objects.Add obj, key
//!         End If
//!     End Sub
//!     
//!     Public Sub UnregisterAll()
//!         Dim i As Long
//!         
//!         For i = m_objects.Count To 1 Step -1
//!             If IsObject(m_objects(i)) Then
//!                 Set m_objects(i) = Nothing
//!             End If
//!         Next i
//!         
//!         Set m_objects = New Collection
//!     End Sub
//!     
//!     Public Function GetStats() As String
//!         Dim msg As String
//!         Dim i As Long
//!         Dim objectCount As Long
//!         Dim nothingCount As Long
//!         
//!         objectCount = 0
//!         nothingCount = 0
//!         
//!         For i = 1 To m_objects.Count
//!             If IsObject(m_objects(i)) Then
//!                 If m_objects(i) Is Nothing Then
//!                     nothingCount = nothingCount + 1
//!                 Else
//!                     objectCount = objectCount + 1
//!                 End If
//!             End If
//!         Next i
//!         
//!         msg = "Total: " & m_objects.Count & vbCrLf
//!         msg = msg & "Active Objects: " & objectCount & vbCrLf
//!         msg = msg & "Nothing: " & nothingCount
//!         
//!         GetStats = msg
//!     End Function
//!     
//!     Private Sub Class_Terminate()
//!         UnregisterAll
//!     End Sub
//! End Class
//!
//! ' Example 3: Variant inspector utility
//! Public Class VariantInspector
//!     Public Function Inspect(value As Variant) As String
//!         Dim info As String
//!         
//!         info = "VarType: " & VarType(value) & vbCrLf
//!         info = info & "TypeName: " & TypeName(value) & vbCrLf
//!         
//!         If IsNull(value) Then
//!             info = info & "State: Null" & vbCrLf
//!         ElseIf IsEmpty(value) Then
//!             info = info & "State: Empty" & vbCrLf
//!         ElseIf IsObject(value) Then
//!             info = info & "State: Object" & vbCrLf
//!             
//!             If value Is Nothing Then
//!                 info = info & "Value: Nothing" & vbCrLf
//!             Else
//!                 info = info & "Value: [" & TypeName(value) & " instance]" & vbCrLf
//!             End If
//!         ElseIf IsArray(value) Then
//!             info = info & "State: Array" & vbCrLf
//!             info = info & "Bounds: " & LBound(value) & " to " & UBound(value) & vbCrLf
//!         Else
//!             info = info & "State: Value" & vbCrLf
//!             info = info & "Value: " & value & vbCrLf
//!         End If
//!         
//!         Inspect = info
//!     End Function
//! End Class
//!
//! ' Example 4: Smart assignment utility
//! Public Class SmartAssignment
//!     Public Sub Assign(ByRef target As Variant, source As Variant)
//!         On Error GoTo ErrorHandler
//!         
//!         If IsObject(source) Then
//!             Set target = source
//!         Else
//!             target = source
//!         End If
//!         
//!         Exit Sub
//!         
//!     ErrorHandler:
//!         Err.Raise Err.Number, "SmartAssignment", _
//!                   "Failed to assign: " & Err.Description
//!     End Sub
//!     
//!     Public Sub CopyArray(ByRef target As Variant, source As Variant)
//!         Dim i As Long
//!         
//!         If Not IsArray(source) Then
//!             Err.Raise 5, , "Source must be an array"
//!         End If
//!         
//!         ReDim target(LBound(source) To UBound(source))
//!         
//!         For i = LBound(source) To UBound(source)
//!             If IsObject(source(i)) Then
//!                 Set target(i) = source(i)
//!             Else
//!                 target(i) = source(i)
//!             End If
//!         Next i
//!     End Sub
//!     
//!     Public Sub ReleaseArray(ByRef arr As Variant)
//!         Dim i As Long
//!         
//!         If Not IsArray(arr) Then Exit Sub
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If IsObject(arr(i)) Then
//!                 Set arr(i) = Nothing
//!             Else
//!                 arr(i) = Empty
//!             End If
//!         Next i
//!         
//!         Erase arr
//!     End Sub
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The IsObject function itself does not raise errors:
//!
//! ```vb
//! ' IsObject is safe to call on any variable
//! Dim obj As Object
//! Dim str As String
//! Dim num As Integer
//!
//! Debug.Print IsObject(obj)      ' True - even if Nothing
//! Debug.Print IsObject(str)      ' False
//! Debug.Print IsObject(num)      ' False
//!
//! ' Common pattern: check before object operations
//! Sub ProcessParameter(param As Variant)
//!     If Not IsObject(param) Then
//!         Err.Raise 5, , "Object required"
//!     End If
//!     
//!     If param Is Nothing Then
//!         Err.Raise 91, , "Object not set"
//!     End If
//!     
//!     ' Safe to use param as object
//!     param.DoSomething
//! End Sub
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: IsObject is a very fast type check
//! - **Variant Overhead**: Using Variant variables has more overhead than typed variables
//! - **Early Binding**: When possible, use typed object variables for better performance
//! - **Frequent Checks**: Cache results if checking same variable multiple times
//!
//! ## Best Practices
//!
//! 1. **Check Before Use**: Use IsObject before accessing object members
//! 2. **Proper Assignment**: Use Set for object assignment, regular = for values
//! 3. **Cleanup Objects**: Set object variables to Nothing when done
//! 4. **Combine with Is Nothing**: Check both IsObject and Is Nothing for complete validation
//! 5. **Type-Specific Checks**: Use TypeOf...Is for specific object type testing
//! 6. **Error Handling**: Provide clear error messages for type mismatches
//! 7. **Variant Collections**: Use IsObject when working with mixed-type collections
//! 8. **Documentation**: Clearly document when functions accept/return objects
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | IsObject | Check if object type | Boolean | Detect object variables |
//! | TypeOf...Is | Check specific object type | Boolean | Test for specific class/interface |
//! | Is Nothing | Check if object is Nothing | Boolean | Test if object is initialized |
//! | VarType | Get variant type | Integer | Detailed type information |
//! | TypeName | Get type name | String | Type name as string |
//! | IsNull | Check if Null | Boolean | Detect Null values |
//! | IsEmpty | Check if uninitialized | Boolean | Detect Empty Variants |
//!
//! ## IsObject vs TypeOf...Is
//!
//! ```vb
//! Dim obj As Object
//! Dim frm As Form
//!
//! Set obj = New Collection
//! Set frm = New Form1
//!
//! ' IsObject checks if variable is any object type
//! Debug.Print IsObject(obj)           ' True
//! Debug.Print IsObject(frm)           ' True
//!
//! ' TypeOf...Is checks for specific object type
//! Debug.Print TypeOf obj Is Collection    ' True
//! Debug.Print TypeOf obj Is Form          ' False
//! Debug.Print TypeOf frm Is Form          ' True
//! Debug.Print TypeOf frm Is Collection    ' False
//!
//! ' Use IsObject for general object detection
//! ' Use TypeOf...Is for specific type validation
//! ```
//!
//! ## IsObject with Nothing
//!
//! ```vb
//! Dim obj As Object
//!
//! ' Object variable not set
//! Debug.Print IsObject(obj)       ' True - it's an object type
//! Debug.Print obj Is Nothing      ' True - but not initialized
//!
//! ' Set to Nothing explicitly
//! Set obj = Nothing
//! Debug.Print IsObject(obj)       ' True - still object type
//! Debug.Print obj Is Nothing      ' True
//!
//! ' Proper validation pattern
//! If IsObject(obj) Then
//!     If Not obj Is Nothing Then
//!         ' Safe to use obj
//!         obj.DoSomething
//!     Else
//!         ' Object type but not initialized
//!         MsgBox "Object not set"
//!     End If
//! Else
//!     ' Not an object type
//!     MsgBox "Not an object"
//! End If
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns Boolean type
//! - Works with all object types (Form, Control, Collection, COM objects, etc.)
//! - Returns True for Nothing (object type, but not initialized)
//! - Critical for proper object lifecycle management
//!
//! ## Limitations
//!
//! - Cannot distinguish between different object types (use TypeOf...Is)
//! - Returns True for Nothing (need to check Is Nothing separately)
//! - Cannot detect if object has been destroyed/released externally
//! - Does not validate object's internal state or usability
//! - Cannot determine object's capabilities or supported interfaces
//!
//! ## Related Functions
//!
//! - `TypeOf...Is`: Check if object is specific type
//! - `Is Nothing`: Check if object reference is Nothing
//! - `VarType`: Get detailed Variant type information (vbObject = 9)
//! - `TypeName`: Get type name as string
//! - `IsNull`: Check if Variant is Null
//! - `IsEmpty`: Check if Variant is Empty

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_isobject_basic() {
        let source = r#"
Sub Test()
    result = IsObject(myVariable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_if_statement() {
        let source = r#"
Sub Test()
    If IsObject(value) Then
        Set result = value
    Else
        result = value
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_not_condition() {
        let source = r#"
Sub Test()
    If Not IsObject(param) Then
        Err.Raise 5, , "Object required"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_function_return() {
        let source = r#"
Function IsAnObject(v As Variant) As Boolean
    IsAnObject = IsObject(v)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_boolean_and() {
        let source = r#"
Sub Test()
    If IsObject(obj) And Not obj Is Nothing Then
        obj.DoSomething
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_boolean_or() {
        let source = r#"
Sub Test()
    If IsObject(field) Or IsNull(field) Then
        ShowWarning
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_iif() {
        let source = r#"
Sub Test()
    result = IIf(IsObject(value), "Object", "Value")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is object: " & IsObject(testVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Object status: " & IsObject(myObj)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_do_while() {
        let source = r#"
Sub Test()
    Do While IsObject(currentItem)
        Set currentItem = GetNext()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_do_until() {
        let source = r#"
Sub Test()
    Do Until Not IsObject(result)
        result = ProcessNext()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_variable_assignment() {
        let source = r#"
Sub Test()
    Dim isObj As Boolean
    isObj = IsObject(dataValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_property_assignment() {
        let source = r#"
Sub Test()
    obj.IsObjectType = IsObject(obj.Value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_isObject = IsObject(m_data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_with_statement() {
        let source = r#"
Sub Test()
    With container
        .HasObject = IsObject(.Item)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_function_argument() {
        let source = r#"
Sub Test()
    Call ValidateType(IsObject(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_select_case() {
        let source = r#"
Sub Test()
    Select Case True
        Case IsObject(value)
            ProcessObject value
        Case Else
            ProcessValue value
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(arr)
        If IsObject(arr(i)) Then
            Set arr(i) = Nothing
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_elseif() {
        let source = r#"
Sub Test()
    If IsNumeric(data) Then
        ProcessNumber data
    ElseIf IsObject(data) Then
        ProcessObject data
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_concatenation() {
        let source = r#"
Sub Test()
    status = "Type: " & IsObject(variable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_parentheses() {
        let source = r#"
Sub Test()
    result = (IsObject(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_array_check() {
        let source = r#"
Sub Test()
    checks(i) = IsObject(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_collection_add() {
        let source = r#"
Sub Test()
    objectFlags.Add IsObject(items(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_comparison() {
        let source = r#"
Sub Test()
    If IsObject(var1) = IsObject(var2) Then
        MsgBox "Same category"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IsObject(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_while_wend() {
        let source = r#"
Sub Test()
    While IsObject(current)
        Set current = GetNextObject()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isobject_cleanup() {
        let source = r#"
Sub Cleanup(ByRef obj As Variant)
    If IsObject(obj) Then
        Set obj = Nothing
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsObject"));
        assert!(text.contains("Identifier"));
    }
}
