//! # `CVErr` Function
//!
//! Returns a `Variant` of subtype Error containing an error number.
//!
//! ## Syntax
//!
//! ```vb
//! CVErr(errornumber)
//! ```
//!
//! ## Parameters
//!
//! - **`errornumber`**: Required. `Long` integer that identifies an error. The valid range is from
//!   0 to 65535, though application-defined errors are typically in the range 513-65535 (VB6
//!   uses 1-512 for system errors).
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `Error` (`VarType = 10`) containing the specified error number.
//! This is not the same as raising an error with `Err.Raise`; instead, it creates an `Error`
//! value that can be assigned to variables and returned from functions.
//!
//! ## Remarks
//!
//! The `CVErr` function is used to create user-defined error values that can be returned from
//! functions or assigned to `Variant` variables. This is particularly useful for:
//!
//! - Returning error conditions from functions without raising exceptions
//! - Creating functions that behave like Excel worksheet functions (returning error values)
//! - Signaling invalid results that should propagate through calculations
//! - Implementing error handling in data processing pipelines
//!
//! **Important Characteristics:**
//!
//! - Returns a `Variant` of subtype `Error` (not an exception)
//! - `Error` values propagate through expressions
//! - Can be tested with `IsError()` function
//! - Not the same as `Err` object or `Err.Raise`
//! - Commonly used with VBA functions called from Excel
//! - `Error` values cannot be used in arithmetic operations
//! - `VarType` of `CVErr` result is 10 (vbError)
//!
//! ## Error Number Ranges
//!
//! | Range | Description |
//! |-------|-------------|
//! | 0-512 | Reserved for VB6 system errors |
//! | 513-65535 | Available for application-defined errors |
//! | 2007-2042 | Excel error values (when used in Excel automation) |
//!
//! ## Excel Error Constants
//!
//! When creating functions for Excel, these error values are commonly used:
//!
//! | Constant | Value | Excel Display | Meaning |
//! |----------|-------|---------------|---------|
//! | xlErrDiv0 | 2007 | #DIV/0! | Division by zero |
//! | xlErrNA | 2042 | #N/A | Value not available |
//! | xlErrName | 2029 | #NAME? | Invalid name |
//! | xlErrNull | 2000 | #NULL! | Null intersection |
//! | xlErrNum | 2036 | #NUM! | Invalid number |
//! | xlErrRef | 2023 | #REF! | Invalid reference |
//! | xlErrValue | 2015 | #VALUE! | Wrong type |
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Create an error value
//! Dim result As Variant
//! result = CVErr(2042)  ' Create #N/A error
//!
//! ' Test if value is an error
//! If IsError(result) Then
//!     MsgBox "Result is an error"
//! End If
//! ```
//!
//! ### Function Returning Error on Invalid Input
//!
//! ```vb
//! Function SafeDivide(numerator As Double, denominator As Double) As Variant
//!     If denominator = 0 Then
//!         SafeDivide = CVErr(2007)  ' #DIV/0!
//!     Else
//!         SafeDivide = numerator / denominator
//!     End If
//! End Function
//!
//! ' Usage
//! Dim result As Variant
//! result = SafeDivide(10, 0)
//! If IsError(result) Then
//!     MsgBox "Division error occurred"
//! Else
//!     MsgBox "Result: " & result
//! End If
//! ```
//!
//! ### Excel UDF with Error Handling
//!
//! ```vb
//! Function Lookup(value As Variant, table As Range) As Variant
//!     Dim cell As Range
//!     
//!     ' Validate input
//!     If IsEmpty(value) Then
//!         Lookup = CVErr(2042)  ' #N/A
//!         Exit Function
//!     End If
//!     
//!     ' Search for value
//!     For Each cell In table
//!         If cell.Value = value Then
//!             Lookup = cell.Offset(0, 1).Value
//!             Exit Function
//!         End If
//!     Next cell
//!     
//!     ' Not found
//!     Lookup = CVErr(2042)  ' #N/A
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Error Constants Definition
//!
//! ```vb
//! ' Define Excel error constants
//! Public Const xlErrDiv0 As Long = 2007   ' #DIV/0!
//! Public Const xlErrNA As Long = 2042     ' #N/A
//! Public Const xlErrName As Long = 2029   ' #NAME?
//! Public Const xlErrNull As Long = 2000   ' #NULL!
//! Public Const xlErrNum As Long = 2036    ' #NUM!
//! Public Const xlErrRef As Long = 2023    ' #REF!
//! Public Const xlErrValue As Long = 2015  ' #VALUE!
//!
//! ' Use in functions
//! Function MyFunction(input As Variant) As Variant
//!     If Not IsNumeric(input) Then
//!         MyFunction = CVErr(xlErrValue)
//!     Else
//!         MyFunction = input * 2
//!     End If
//! End Function
//! ```
//!
//! ### Validation Functions
//!
//! ```vb
//! Function ValidateRange(value As Double, min As Double, max As Double) As Variant
//!     If value < min Or value > max Then
//!         ValidateRange = CVErr(2036)  ' #NUM! - Number out of range
//!     Else
//!         ValidateRange = value
//!     End If
//! End Function
//! ```
//!
//! ### Error Propagation
//!
//! ```vb
//! Function Calculate(input As Variant) As Variant
//!     ' Check if input is already an error
//!     If IsError(input) Then
//!         Calculate = input  ' Propagate the error
//!         Exit Function
//!     End If
//!     
//!     ' Perform calculation
//!     If input < 0 Then
//!         Calculate = CVErr(2036)  ' #NUM!
//!     Else
//!         Calculate = Sqr(input)
//!     End If
//! End Function
//! ```
//!
//! ### Database Lookup with Error
//!
//! ```vb
//! Function GetEmployeeName(employeeID As Long) As Variant
//!     Dim rs As ADODB.Recordset
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     Set rs = New ADODB.Recordset
//!     rs.Open "SELECT Name FROM Employees WHERE ID = " & employeeID, conn
//!     
//!     If rs.EOF Then
//!         GetEmployeeName = CVErr(2042)  ' #N/A - Not found
//!     Else
//!         GetEmployeeName = rs("Name")
//!     End If
//!     
//!     rs.Close
//!     Set rs = Nothing
//!     Exit Function
//!     
//! ErrorHandler:
//!     GetEmployeeName = CVErr(2042)  ' #N/A
//! End Function
//! ```
//!
//! ### Array Formula with Errors
//!
//! ```vb
//! Function ProcessArray(values As Variant) As Variant
//!     Dim results() As Variant
//!     Dim i As Long
//!     
//!     If Not IsArray(values) Then
//!         ProcessArray = CVErr(2015)  ' #VALUE!
//!         Exit Function
//!     End If
//!     
//!     ReDim results(LBound(values) To UBound(values))
//!     
//!     For i = LBound(values) To UBound(values)
//!         If IsError(values(i)) Then
//!             results(i) = values(i)  ' Propagate error
//!         ElseIf Not IsNumeric(values(i)) Then
//!             results(i) = CVErr(2015)  ' #VALUE!
//!         Else
//!             results(i) = values(i) * 2
//!         End If
//!     Next i
//!     
//!     ProcessArray = results
//! End Function
//! ```
//!
//! ### Type Conversion with Error
//!
//! ```vb
//! Function SafeCLng(value As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     If IsError(value) Then
//!         SafeCLng = value  ' Propagate error
//!     ElseIf IsEmpty(value) Then
//!         SafeCLng = CVErr(2042)  ' #N/A
//!     ElseIf Not IsNumeric(value) Then
//!         SafeCLng = CVErr(2015)  ' #VALUE!
//!     Else
//!         Dim temp As Double
//!         temp = CDbl(value)
//!         If temp < -2147483648# Or temp > 2147483647# Then
//!             SafeCLng = CVErr(2036)  ' #NUM!
//!         Else
//!             SafeCLng = CLng(value)
//!         End If
//!     End If
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeCLng = CVErr(2015)  ' #VALUE!
//! End Function
//! ```
//!
//! ### Conditional Error Return
//!
//! ```vb
//! Function GetDiscount(totalSales As Double) As Variant
//!     Select Case totalSales
//!         Case Is < 0
//!             GetDiscount = CVErr(2036)  ' #NUM! - Negative sales
//!         Case 0 To 999.99
//!             GetDiscount = 0
//!         Case 1000 To 4999.99
//!             GetDiscount = 0.05
//!         Case 5000 To 9999.99
//!             GetDiscount = 0.1
//!         Case Is >= 10000
//!             GetDiscount = 0.15
//!         Case Else
//!             GetDiscount = CVErr(2042)  ' #N/A
//!     End Select
//! End Function
//! ```
//!
//! ### Error Checking Helper
//!
//! ```vb
//! Function GetErrorNumber(value As Variant) As Long
//!     ' Returns error number or 0 if not an error
//!     If IsError(value) Then
//!         ' There's no direct way to extract error number in VB6
//!         ' Would need to compare against known error values
//!         GetErrorNumber = -1  ' Indicates error present
//!     Else
//!         GetErrorNumber = 0
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Custom Error Type System
//!
//! ```vb
//! ' Application-specific error codes
//! Public Const APP_ERR_INVALID_USER As Long = 1000
//! Public Const APP_ERR_DATABASE As Long = 1001
//! Public Const APP_ERR_NETWORK As Long = 1002
//! Public Const APP_ERR_PERMISSION As Long = 1003
//!
//! Function AuthenticateUser(username As String, password As String) As Variant
//!     If Len(username) = 0 Then
//!         AuthenticateUser = CVErr(APP_ERR_INVALID_USER)
//!         Exit Function
//!     End If
//!     
//!     ' Check credentials
//!     If Not ValidateCredentials(username, password) Then
//!         AuthenticateUser = CVErr(APP_ERR_PERMISSION)
//!         Exit Function
//!     End If
//!     
//!     ' Return user object on success
//!     AuthenticateUser = GetUserObject(username)
//! End Function
//! ```
//!
//! ### Error Value in Collections
//!
//! ```vb
//! Function ProcessRecords() As Collection
//!     Dim results As New Collection
//!     Dim rs As ADODB.Recordset
//!     Dim i As Long
//!     
//!     Set rs = GetRecords()
//!     
//!     While Not rs.EOF
//!         On Error Resume Next
//!         results.Add ProcessRecord(rs)
//!         
//!         If Err.Number <> 0 Then
//!             results.Add CVErr(2015)  ' Add error marker
//!             Err.Clear
//!         End If
//!         
//!         rs.MoveNext
//!     Wend
//!     
//!     Set ProcessRecords = results
//! End Function
//! ```
//!
//! ### Chainable Operations with Error Propagation
//!
//! ```vb
//! Function Step1(input As Variant) As Variant
//!     If IsError(input) Then
//!         Step1 = input
//!     ElseIf input < 0 Then
//!         Step1 = CVErr(2036)
//!     Else
//!         Step1 = Sqr(input)
//!     End If
//! End Function
//!
//! Function Step2(input As Variant) As Variant
//!     If IsError(input) Then
//!         Step2 = input
//!     ElseIf input = 0 Then
//!         Step2 = CVErr(2007)
//!     Else
//!         Step2 = 100 / input
//!     End If
//! End Function
//!
//! ' Chain operations
//! result = Step2(Step1(value))
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function CreateErrorSafe(errorNum As Long) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     If errorNum < 0 Or errorNum > 65535 Then
//!         CreateErrorSafe = CVErr(2036)  ' Invalid error number
//!     Else
//!         CreateErrorSafe = CVErr(errorNum)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     CreateErrorSafe = CVErr(2042)  ' Generic error
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type mismatch): Error number is not a valid Long integer
//! - **Error 6** (Overflow): Error number is outside valid range (0-65535)
//!
//! ## Performance Considerations
//!
//! - `CVErr` is a fast function with minimal overhead
//! - `Error` values are lightweight `Variant` subtypes
//! - Using `CVErr` is more efficient than raising and catching exceptions
//! - `Error` propagation through calculations is automatic
//! - No significant memory overhead compared to other `Variant` values
//!
//! ## Comparison with Other Error Mechanisms
//!
//! ### `CVErr` vs `Err.Raise`
//!
//! ```vb
//! ' CVErr - Returns error value (doesn't stop execution)
//! Function Method1(x As Double) As Variant
//!     If x < 0 Then
//!         Method1 = CVErr(2036)  ' Returns error, continues
//!     Else
//!         Method1 = Sqr(x)
//!     End If
//! End Function
//!
//! ' Err.Raise - Throws exception (stops execution)
//! Function Method2(x As Double) As Double
//!     If x < 0 Then
//!         Err.Raise 5, , "Invalid argument"  ' Stops execution
//!     Else
//!         Method2 = Sqr(x)
//!     End If
//! End Function
//! ```
//!
//! **`CVErr` advantages:**
//! - Doesn't interrupt program flow
//! - Can be used in expressions
//! - Natural for functional-style programming
//! - Compatible with Excel worksheet functions
//!
//! **`Err.Raise` advantages:**
//! - Forces immediate attention to errors
//! - Provides error description and source
//! - Traditional exception handling model
//! - Better for critical errors
//!
//! ## Best Practices
//!
//! ### Always Check for Errors Before Using Values
//!
//! ```vb
//! Dim result As Variant
//! result = SomeFunction()
//!
//! If IsError(result) Then
//!     MsgBox "Error occurred"
//! Else
//!     ' Safe to use result
//!     Debug.Print result
//! End If
//! ```
//!
//! ### Use Meaningful Error Numbers
//!
//! ```vb
//! ' Good - Use named constants
//! Const ERR_INVALID_INPUT As Long = 1000
//! result = CVErr(ERR_INVALID_INPUT)
//!
//! ' Avoid - Magic numbers
//! result = CVErr(42)  ' What does 42 mean?
//! ```
//!
//! ### Document Custom Error Codes
//!
//! ```vb
//! ' Application Error Codes (1000-1999)
//! Public Const ERR_INVALID_USER As Long = 1000    ' Invalid username
//! Public Const ERR_EXPIRED_SESSION As Long = 1001  ' Session expired
//! Public Const ERR_INSUFFICIENT_RIGHTS As Long = 1002  ' Access denied
//! ```
//!
//! ## Limitations
//!
//! - Cannot extract error number from error value directly in VB6
//! - `Error` values cannot be used in arithmetic operations
//! - Limited to Long integer error numbers (0-65535)
//! - No built-in error description with `CVErr` (unlike `Err` object)
//! - `VarType` test required to detect errors (`IsError` function)
//! - Not all VB6 functions handle error values gracefully
//!
//! ## Related Functions
//!
//! - `IsError`: Tests if a Variant contains an error value
//! - `Err.Raise`: Raises a runtime error (different from `CVErr`)
//! - `Error`: Returns error message for an error number
//! - `Error$`: `String` version of `Error` function
//! - `VarType`: Returns the subtype of a Variant (10 for `Error`)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn cverr_basic() {
        let source = r"
result = CVErr(2042)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2042"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_division_by_zero() {
        let source = r"
error = CVErr(2007)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            ErrorStatement {
                ErrorKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("CVErr"),
                LeftParenthesis,
                IntegerLiteral ("2007"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn cverr_in_function() {
        let source = r"
Function SafeDivide(a As Double, b As Double) As Variant
    If b = 0 Then
        SafeDivide = CVErr(2007)
    Else
        SafeDivide = a / b
    End If
End Function
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("SafeDivide"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("a"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("b"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("b"),
                            },
                            Whitespace,
                            EqualityOperator,
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
                                    Identifier ("SafeDivide"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("CVErr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2007"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("SafeDivide"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("a"),
                                        },
                                        Whitespace,
                                        DivisionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn cverr_with_constant() {
        let source = r"
Const xlErrNA As Long = 2042
result = CVErr(xlErrNA)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            DimStatement {
                ConstKeyword,
                Whitespace,
                Identifier ("xlErrNA"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("2042"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("xlErrNA"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_value_error() {
        let source = r"
err = CVErr(2015)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("err"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2015"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_with_iserror() {
        let source = r#"
result = CVErr(2042)
If IsError(result) Then
    MsgBox "Error"
End If
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2042"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            IfStatement {
                IfKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("IsError"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("result"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_in_assignment() {
        let source = r"
Dim myError As Variant
myError = CVErr(2036)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("myError"),
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("myError"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2036"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_custom_error() {
        let source = r"
Const APP_ERR_INVALID As Long = 1000
result = CVErr(APP_ERR_INVALID)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            DimStatement {
                ConstKeyword,
                Whitespace,
                Identifier ("APP_ERR_INVALID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("1000"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("APP_ERR_INVALID"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_ref_error() {
        let source = r"
refError = CVErr(2023)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("refError"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2023"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_num_error() {
        let source = r"
numError = CVErr(2036)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("numError"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2036"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_in_select_case() {
        let source = r"
Select Case value
    Case Is < 0
        result = CVErr(2036)
    Case Else
        result = value
End Select
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("value"),
                },
                Newline,
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    IsKeyword,
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    IntegerLiteral ("0"),
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
                                Identifier ("CVErr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("2036"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseElseClause {
                    CaseKeyword,
                    Whitespace,
                    ElseKeyword,
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
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                SelectKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_propagation() {
        let source = r"
If IsError(input) Then
    output = input
Else
    output = CVErr(2042)
End If
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("IsError"),
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
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            OutputKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            InputKeyword,
                        },
                        Newline,
                    },
                },
                ElseClause {
                    ElseKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                OutputKeyword,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("CVErr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("2042"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_name_error() {
        let source = r"
nameErr = CVErr(2029)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("nameErr"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2029"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_null_error() {
        let source = r"
nullErr = CVErr(2000)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("nullErr"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2000"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_in_array() {
        let source = r"
results(i) = CVErr(2042)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                CallExpression {
                    Identifier ("results"),
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
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2042"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_with_variable() {
        let source = r"
Dim errorNum As Long
errorNum = 2042
result = CVErr(errorNum)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("errorNum"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("errorNum"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("2042"),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("errorNum"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_in_if_condition() {
        let source = r"
If value < 0 Then
    result = CVErr(2036)
End If
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("value"),
                    },
                    Whitespace,
                    LessThanOperator,
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
                        CallExpression {
                            Identifier ("CVErr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2036"),
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
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_multiple_errors() {
        let source = r"
err1 = CVErr(2007)
err2 = CVErr(2042)
err3 = CVErr(2015)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("err1"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2007"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("err2"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2042"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("err3"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2015"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn cverr_in_loop() {
        let source = r"
For i = 1 To 10
    If arr(i) < 0 Then
        arr(i) = CVErr(2036)
    End If
Next i
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            ForStatement {
                ForKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("i"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("1"),
                },
                Whitespace,
                ToKeyword,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("10"),
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
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
                            LessThanOperator,
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
                                    Identifier ("CVErr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2036"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
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
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_return_value() {
        let source = r"
Function Validate(x As Variant) As Variant
    Validate = CVErr(2015)
End Function
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Validate"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    VariantKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Validate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("CVErr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2015"),
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
    fn cverr_in_do_loop() {
        let source = r#"
Do While Not rs.EOF
    If IsNull(rs("Value")) Then
        result = CVErr(2042)
    End If
    rs.MoveNext
Loop
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                UnaryExpression {
                    NotKeyword,
                    Whitespace,
                    MemberAccessExpression {
                        Identifier ("rs"),
                        PeriodOperator,
                        Identifier ("EOF"),
                    },
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("IsNull"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("rs"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"Value\""),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
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
                                CallExpression {
                                    Identifier ("CVErr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2042"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
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
                    Whitespace,
                    CallStatement {
                        Identifier ("rs"),
                        PeriodOperator,
                        Identifier ("MoveNext"),
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_with_expression() {
        let source = r"
errCode = baseError + offset
result = CVErr(errCode)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("errCode"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("baseError"),
                    },
                    Whitespace,
                    AdditionOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("offset"),
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("errCode"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_in_collection() {
        let source = r"
errors.Add CVErr(2042)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("errors"),
                PeriodOperator,
                Identifier ("Add"),
                Whitespace,
                Identifier ("CVErr"),
                LeftParenthesis,
                IntegerLiteral ("2042"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_zero() {
        let source = r"
err = CVErr(0)
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("err"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn cverr_with_whitespace() {
        let source = r"
result = CVErr( 2042 )
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("CVErr"),
                    LeftParenthesis,
                    ArgumentList {
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2042"),
                            },
                        },
                        Whitespace,
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }
}
