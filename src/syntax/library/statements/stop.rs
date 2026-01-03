//! # Stop Statement
//!
//! Suspends execution.
//!
//! ## Syntax
//!
//! ```vb
//! Stop
//! ```
//!
//! ## Remarks
//!
//! - **Execution Suspension**: The Stop statement suspends execution but doesn't close any files or clear variables unless it is in a compiled executable (.exe) file.
//! - **Debug Mode**: In the development environment, Stop causes the program to enter break mode, allowing you to examine variables, step through code, and use debugging tools.
//! - **Compiled Executable**: In a compiled .exe file, Stop acts like the End statement, terminating the program and closing all files.
//! - **Break Mode**: When Stop is encountered in the IDE, VB6 pauses execution and highlights the Stop statement, allowing you to inspect the current state.
//! - **Multiple Stop Statements**: You can place Stop statements anywhere in your code to create breakpoints for debugging.
//! - **Not for Production**: Stop statements should generally be removed before distributing your application, as they can cause unexpected behavior.
//! - **Alternative to Breakpoints**: Stop provides a code-based alternative to setting breakpoints in the IDE.
//! - **No Arguments**: The Stop statement takes no arguments or parameters.
//!
//! ## Common Uses
//!
//! - **Debugging**: Pause execution to examine variable values and program state
//! - **Conditional Breakpoints**: Combined with If statements for conditional debugging
//! - **Error Investigation**: Stop execution when an error condition is detected
//! - **Loop Debugging**: Pause execution during specific loop iterations
//! - **Testing**: Verify code paths during development
//!
//! ## Examples
//!
//! ### Simple Stop
//!
//! ```vb
//! Sub Test()
//!     Dim x As Integer
//!     x = 10
//!     Stop  ' Execution pauses here in IDE
//!     x = x + 5
//! End Sub
//! ```
//!
//! ### Conditional Stop for Debugging
//!
//! ```vb
//! Sub ProcessData(value As Integer)
//!     If value < 0 Then
//!         Stop  ' Pause when invalid data is encountered
//!     End If
//!     ' Process value
//! End Sub
//! ```
//!
//! ### Stop in Loop for Specific Iteration
//!
//! ```vb
//! For i = 1 To 100
//!     If i = 50 Then
//!         Stop  ' Pause at iteration 50
//!     End If
//!     ProcessItem i
//! Next i
//! ```
//!
//! ### Stop on Error Condition
//!
//! ```vb
//! Sub CalculateTotal()
//!     Dim total As Double
//!     total = GetSubtotal()
//!     
//!     If total < 0 Then
//!         Stop  ' Investigate negative total
//!     End If
//!     
//!     SaveTotal total
//! End Sub
//! ```
//!
//! ### Multiple Stop Statements for Debugging Path
//!
//! ```vb
//! Function ValidateData(data As String) As Boolean
//!     Stop  ' Entry point
//!     
//!     If Len(data) = 0 Then
//!         Stop  ' Empty string case
//!         ValidateData = False
//!         Exit Function
//!     End If
//!     
//!     Stop  ' Normal processing
//!     ValidateData = True
//! End Function
//! ```
//!
//! ### Stop in Select Case
//!
//! ```vb
//! Select Case userType
//!     Case 1
//!         ProcessAdmin
//!     Case 2
//!         ProcessUser
//!     Case Else
//!         Stop  ' Unknown user type - investigate
//! End Select
//! ```
//!
//! ### Stop with Error Handler
//!
//! ```vb
//! On Error GoTo ErrorHandler
//!
//! ProcessData
//! Exit Sub
//!
//! ErrorHandler:
//!     Stop  ' Pause to examine error
//!     MsgBox Err.Description
//! End Sub
//! ```
//!
//! ### Stop in Class Module
//!
//! ```vb
//! Private Sub Class_Initialize()
//!     Stop  ' Verify initialization sequence
//!     InitializeProperties
//! End Sub
//! ```
//!
//! ### Stop Before Critical Operation
//!
//! ```vb
//! Sub DeleteAllRecords()
//!     Stop  ' Verify this operation should proceed
//!     
//!     Dim rs As Recordset
//!     Set rs = db.OpenRecordset("Data")
//!     
//!     Do While Not rs.EOF
//!         rs.Delete
//!         rs.MoveNext
//!     Loop
//! End Sub
//! ```
//!
//! ### Stop in Property Procedure
//!
//! ```vb
//! Public Property Let Value(ByVal newValue As Integer)
//!     If newValue < 0 Then
//!         Stop  ' Negative value assigned
//!     End If
//!     m_Value = newValue
//! End Property
//! ```
//!
//! ### Stop with `DoEvents`
//!
//! ```vb
//! For i = 1 To 1000
//!     DoEvents
//!     ProcessItem i
//!     
//!     If ShouldDebug Then
//!         Stop  ' Conditional pause
//!     End If
//! Next i
//! ```
//!
//! ### Stop in Event Handler
//!
//! ```vb
//! Private Sub Form_Load()
//!     Stop  ' Debug form initialization
//!     InitializeControls
//!     LoadData
//! End Sub
//! ```
//!
//! ### Stop with Assert-like Check
//!
//! ```vb
//! Sub ProcessArray(arr() As Integer)
//!     If UBound(arr) < LBound(arr) Then
//!         Stop  ' Invalid array bounds
//!     End If
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         ProcessItem arr(i)
//!     Next i
//! End Sub
//! ```
//!
//! ## Important Notes
//!
//! - **IDE vs. Compiled**: Behavior differs between development environment and compiled executables
//! - **Production Code**: Remove Stop statements before distributing your application
//! - **Alternative Debugging**: Modern IDEs prefer breakpoints over Stop statements
//! - **No File Closure**: In IDE, Stop doesn't close files or clear variables
//! - **Break Mode**: Allows interactive debugging in the IDE
//! - **End Alternative**: In compiled .exe, behaves like the End statement
//! - **Code-Based Breakpoint**: Useful when you need a breakpoint that travels with the code
//! - **No Performance Impact**: When compiled, can be configured to be removed by compiler
//!
//! ## Best Practices
//!
//! - Use Stop for temporary debugging during development
//! - Remove Stop statements before final release
//! - Use meaningful comments explaining why Stop is placed at a location
//! - Consider using conditional compilation to automatically remove Stop in release builds
//! - Prefer IDE breakpoints for most debugging scenarios
//! - Use Stop when you need to share debugging points with team members
//! - Document any Stop statements that remain for legitimate reasons
//!
//! ## Differences from End
//!
//! - **Stop**: In IDE, enters break mode; in .exe, terminates program
//! - **End**: Always terminates program and closes all files
//! - **Exit**: Exits specific procedure/function/loop without terminating program
//!
//! ## See Also
//!
//! - `End` statement (terminate program)
//! - `Exit` statement (exit procedure, function, or loop)
//! - `DoEvents` function (yield execution to the operating system)
//! - `Debug.Assert` method (conditional debugging)
//!
//! ## References
//!
//! - [Stop Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/stop-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Stop statement.
    pub(crate) fn parse_stop_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::StopStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn stop_simple() {
        let source = r"
Sub Test()
    Stop
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
                    StopStatement {
                        Whitespace,
                        StopKeyword,
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
    fn stop_at_module_level() {
        let source = "Stop\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [StopStatement {
                StopKeyword,
                Newline,
            },]
        );
        let debug = cst.debug_tree();
        assert!(debug.contains("StopStatement"));
    }

    #[test]
    fn stop_with_comment() {
        let source = r"
Sub Test()
    Stop ' Debug breakpoint
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
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Whitespace,
                        EndOfLineComment,
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
    fn stop_in_if_statement() {
        let source = r"
If value < 0 Then
    Stop
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
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
                    StopStatement {
                        Whitespace,
                        StopKeyword,
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
    fn stop_in_loop() {
        let source = r"
For i = 1 To 100
    If i = 50 Then
        Stop
    End If
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
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
                    IntegerLiteral ("100"),
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("50"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
    fn stop_in_error_handler() {
        let source = r"
On Error GoTo ErrorHandler
ProcessData
Exit Sub

ErrorHandler:
    Stop
    MsgBox Err.Description
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                GotoKeyword,
                Whitespace,
                Identifier ("ErrorHandler"),
                Newline,
            },
            CallStatement {
                Identifier ("ProcessData"),
                Newline,
            },
            ExitStatement {
                ExitKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
            Newline,
            LabelStatement {
                Identifier ("ErrorHandler"),
                ColonOperator,
                Newline,
            },
            Whitespace,
            StopStatement {
                StopKeyword,
                Newline,
            },
            Whitespace,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                Identifier ("Err"),
                PeriodOperator,
                Identifier ("Description"),
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_preserves_whitespace() {
        let source = "    Stop    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [
                Whitespace,
                StopStatement {
                    StopKeyword,
                    Whitespace,
                    Newline,
                },
            ]
        );
    }

    #[test]
    fn stop_multiple_on_same_line() {
        let source = "Stop: Stop\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [StopStatement {
                StopKeyword,
                ColonOperator,
                Whitespace,
                StopKeyword,
                Newline,
            },]
        );
    }

    #[test]
    fn stop_in_select_case() {
        let source = r"
Select Case userType
    Case 1
        ProcessAdmin
    Case Else
        Stop
End Select
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("userType"),
                },
                Newline,
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    IntegerLiteral ("1"),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("ProcessAdmin"),
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
                        StopStatement {
                            Whitespace,
                            StopKeyword,
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
    fn stop_in_function() {
        let source = r"
Function ValidateData(data As String) As Boolean
    Stop
    ValidateData = True
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateData"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ValidateData"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
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
    fn stop_in_sub() {
        let source = r"
Sub ProcessData()
    Stop
    ' Process here
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_in_class_initialize() {
        let source = r"
Private Sub Class_Initialize()
    Stop
    InitializeProperties
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Class_Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("InitializeProperties"),
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
    fn stop_before_critical_operation() {
        let source = r"
Sub DeleteAllRecords()
    Stop
    
    Dim rs As Recordset
    rs.Delete
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("DeleteAllRecords"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    Newline,
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("rs"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("Recordset"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("rs"),
                        PeriodOperator,
                        Identifier ("Delete"),
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
    fn stop_in_property_let() {
        let source = r"
Public Property Let Value(ByVal newValue As Integer)
    If newValue < 0 Then
        Stop
    End If
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                LetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("newValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("newValue"),
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
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_with_doevents() {
        let source = r"
For i = 1 To 1000
    DoEvents
    If ShouldDebug Then
        Stop
    End If
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
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
                    IntegerLiteral ("1000"),
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("DoEvents"),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("ShouldDebug"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
    fn stop_in_event_handler() {
        let source = r"
Private Sub Form_Load()
    Stop
    InitializeControls
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Form_Load"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("InitializeControls"),
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
    fn stop_with_array_check() {
        let source = r"
Sub ProcessArray(arr() As Integer)
    If UBound(arr) < LBound(arr) Then
        Stop
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessArray"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("arr"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
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
                            Whitespace,
                            LessThanOperator,
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
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
    fn stop_multiple_in_function() {
        let source = r"
Function ValidateData(data As String) As Boolean
    Stop
    
    If Len(data) = 0 Then
        Stop
        ValidateData = False
        Exit Function
    End If
    
    Stop
    ValidateData = True
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateData"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    Newline,
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                LenKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("data"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
                            StopStatement {
                                Whitespace,
                                StopKeyword,
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("ValidateData"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BooleanLiteralExpression {
                                    FalseKeyword,
                                },
                                Newline,
                            },
                            ExitStatement {
                                Whitespace,
                                ExitKeyword,
                                Whitespace,
                                FunctionKeyword,
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
                    Newline,
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ValidateData"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
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
    fn stop_inline_if() {
        let source = "If debug Then Stop\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            IfStatement {
                IfKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("debug"),
                },
                Whitespace,
                ThenKeyword,
                Whitespace,
                StopStatement {
                    StopKeyword,
                    Newline,
                },
            },
        ]);
    }

    #[test]
    fn stop_in_with_block() {
        let source = r"
With objData
    Stop
    .Property = Value
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("objData"),
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                PropertyKeyword,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("Value"),
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_in_do_loop() {
        let source = r"
Do While Not rs.EOF
    Stop
    ProcessRecord rs
    rs.MoveNext
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
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
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessRecord"),
                        Whitespace,
                        Identifier ("rs"),
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
    fn stop_conditional_debugging() {
        let source = r"
Sub ProcessData(value As Integer)
    If value < 0 Then
        Stop
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
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
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
    fn stop_in_class_terminate() {
        let source = r"
Private Sub Class_Terminate()
    Stop
    Cleanup
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Class_Terminate"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Cleanup"),
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
    fn stop_in_property_get() {
        let source = r"
Public Property Get Value() As Integer
    Stop
    Value = m_Value
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PublicKeyword,
                Whitespace,
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("Value"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList {
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("m_Value"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_after_calculation() {
        let source = r"
Sub Calculate()
    Dim total As Double
    total = GetSubtotal()
    
    If total < 0 Then
        Stop
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Calculate"),
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
                        Identifier ("total"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("total"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetSubtotal"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    Newline,
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("total"),
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
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
    fn stop_in_for_each() {
        let source = r"
For Each item In collection
    If item.IsInvalid Then
        Stop
    End If
    ProcessItem item
Next item
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForEachStatement {
                ForKeyword,
                Whitespace,
                EachKeyword,
                Whitespace,
                Identifier ("item"),
                Whitespace,
                InKeyword,
                Whitespace,
                Identifier ("collection"),
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        MemberAccessExpression {
                            Identifier ("item"),
                            PeriodOperator,
                            Identifier ("IsInvalid"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
                        Identifier ("ProcessItem"),
                        Whitespace,
                        Identifier ("item"),
                        Newline,
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("item"),
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_case_insensitive() {
        let source = "stop\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [StopStatement {
                StopKeyword,
                Newline,
            },]
        );
    }

    #[test]
    fn stop_in_multiline_if() {
        let source = r"
If condition1 Then
    If condition2 Then
        Stop
    End If
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("condition1"),
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("condition2"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            StopStatement {
                                Whitespace,
                                StopKeyword,
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
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn stop_multiple_statements() {
        let source = r"
Sub Test()
    Stop
    ProcessData
    Stop
    SaveData
    Stop
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
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessData"),
                        Newline,
                    },
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("SaveData"),
                        Newline,
                    },
                    StopStatement {
                        Whitespace,
                        StopKeyword,
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
    fn stop_in_type_initialization() {
        let source = r"
Sub InitializeType()
    Dim data As MyType
    Stop
    data.Field1 = 10
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("InitializeType"),
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
                        Identifier ("data"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("MyType"),
                        Newline,
                    },
                    StopStatement {
                        Whitespace,
                        StopKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("data"),
                            PeriodOperator,
                            Identifier ("Field1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
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
