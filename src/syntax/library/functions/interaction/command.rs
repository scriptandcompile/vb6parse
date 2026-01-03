//! # `Command` Function
//!
//! Returns the argument portion of the command line used to launch Microsoft Visual Basic or an
//! executable program developed with Visual Basic.
//!
//! ## Syntax
//!
//! ```vb
//! Command()
//! ```
//!
//! ## Parameters
//!
//! None. The `Command` function takes no arguments.
//!
//! ## Return Value
//!
//! Returns a String containing the command-line arguments passed to the program. If no arguments
//! were passed, returns an empty string ("").
//!
//! ## Remarks
//!
//! The `Command` function provides access to the command-line arguments that were passed when the
//! application was started. This is commonly used for:
//!
//! - Processing startup parameters
//! - Accepting file paths to open
//! - Enabling debug or special modes
//! - Configuring application behavior at launch
//!
//! **Important Characteristics:**
//!
//! - Returns only the arguments, not the executable path
//! - Arguments are returned as a single string
//! - Multiple arguments are separated by spaces (unless quoted)
//! - Quoted strings are preserved but quotes may be included in the result
//! - Leading and trailing spaces are typically trimmed
//! - Returns empty string ("") if no arguments provided
//! - Case is preserved as entered
//!
//! ## Command Line Processing
//!
//! When an application is launched with:
//! ```text
//! MyApp.exe /debug file.txt "long filename.doc"
//! ```
//!
//! `Command()` returns:
//! ```text
//! /debug file.txt "long filename.doc"
//! ```
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get command line arguments
//! Sub Main()
//!     Dim cmdLine As String
//!     cmdLine = Command()
//!     
//!     If cmdLine <> "" Then
//!         MsgBox "Arguments: " & cmdLine
//!     Else
//!         MsgBox "No arguments provided"
//!     End If
//! End Sub
//! ```
//!
//! ### Processing Switches
//!
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     args = Command()
//!     
//!     If InStr(args, "/debug") > 0 Then
//!         App.LogMode = 1  ' Enable debug logging
//!     End If
//!     
//!     If InStr(args, "/silent") > 0 Then
//!         App.SilentMode = True
//!     End If
//! End Sub
//! ```
//!
//! ### Opening a File from Command Line
//!
//! ```vb
//! Sub Main()
//!     Dim filename As String
//!     filename = Trim(Command())
//!     
//!     If filename <> "" Then
//!         ' Remove quotes if present
//!         If Left(filename, 1) = Chr(34) Then
//!             filename = Mid(filename, 2)
//!         End If
//!         If Right(filename, 1) = Chr(34) Then
//!             filename = Left(filename, Len(filename) - 1)
//!         End If
//!         
//!         ' Open the file
//!         If Dir(filename) <> "" Then
//!             OpenDocument filename
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Parsing Multiple Arguments
//!
//! ```vb
//! Function ParseCommandLine() As Collection
//!     Dim args As New Collection
//!     Dim cmdLine As String
//!     Dim arg As String
//!     Dim pos As Integer
//!     Dim inQuotes As Boolean
//!     Dim i As Integer
//!     Dim ch As String
//!     
//!     cmdLine = Trim(Command())
//!     If cmdLine = "" Then Exit Function
//!     
//!     arg = ""
//!     inQuotes = False
//!     
//!     For i = 1 To Len(cmdLine)
//!         ch = Mid(cmdLine, i, 1)
//!         
//!         If ch = Chr(34) Then  ' Quote character
//!             inQuotes = Not inQuotes
//!         ElseIf ch = " " And Not inQuotes Then
//!             If arg <> "" Then
//!                 args.Add arg
//!                 arg = ""
//!             End If
//!         Else
//!             arg = arg & ch
//!         End If
//!     Next i
//!     
//!     If arg <> "" Then args.Add arg
//!     
//!     Set ParseCommandLine = args
//! End Function
//! ```
//!
//! ### Named Parameters
//!
//! ```vb
//! Function GetParameter(paramName As String) As String
//!     Dim cmdLine As String
//!     Dim pos As Integer
//!     Dim endPos As Integer
//!     Dim value As String
//!     
//!     cmdLine = " " & Command() & " "
//!     pos = InStr(1, cmdLine, "/" & paramName & ":", vbTextCompare)
//!     
//!     If pos = 0 Then
//!         pos = InStr(1, cmdLine, "-" & paramName & ":", vbTextCompare)
//!     End If
//!     
//!     If pos > 0 Then
//!         pos = InStr(pos, cmdLine, ":") + 1
//!         endPos = InStr(pos, cmdLine, " ")
//!         
//!         If endPos > pos Then
//!             value = Mid(cmdLine, pos, endPos - pos)
//!             GetParameter = Trim(value)
//!         End If
//!     End If
//! End Function
//!
//! ' Usage:
//! ' MyApp.exe /server:localhost /port:8080
//! ' server = GetParameter("server")  ' Returns "localhost"
//! ' port = GetParameter("port")      ' Returns "8080"
//! ```
//!
//! ### Switch Detection
//!
//! ```vb
//! Function HasSwitch(switchName As String) As Boolean
//!     Dim cmdLine As String
//!     cmdLine = " " & LCase(Command()) & " "
//!     
//!     HasSwitch = InStr(cmdLine, " /" & LCase(switchName)) > 0 Or _
//!                 InStr(cmdLine, " -" & LCase(switchName)) > 0
//! End Function
//!
//! ' Usage:
//! ' MyApp.exe /debug /verbose
//! ' If HasSwitch("debug") Then ...
//! ```
//!
//! ### File Association Handler
//!
//! ```vb
//! Sub Main()
//!     Dim filename As String
//!     
//!     filename = GetCommandLineFile()
//!     
//!     If filename <> "" Then
//!         ' Application was launched by double-clicking a file
//!         LoadFile filename
//!     Else
//!         ' Application was launched normally
//!         ShowStartupDialog
//!     End If
//! End Sub
//!
//! Function GetCommandLineFile() As String
//!     Dim cmdLine As String
//!     cmdLine = Trim(Command())
//!     
//!     ' Remove surrounding quotes
//!     If Left(cmdLine, 1) = Chr(34) And Right(cmdLine, 1) = Chr(34) Then
//!         cmdLine = Mid(cmdLine, 2, Len(cmdLine) - 2)
//!     End If
//!     
//!     ' Check if it's a file (not a switch)
//!     If Left(cmdLine, 1) <> "/" And Left(cmdLine, 1) <> "-" Then
//!         If Dir(cmdLine) <> "" Then
//!             GetCommandLineFile = cmdLine
//!         End If
//!     End If
//! End Function
//! ```
//!
//! ### Configuration File Loading
//!
//! ```vb
//! Sub Main()
//!     Dim configFile As String
//!     
//!     configFile = GetParameter("config")
//!     
//!     If configFile = "" Then
//!         configFile = App.Path & "\default.cfg"
//!     End If
//!     
//!     LoadConfiguration configFile
//! End Sub
//! ```
//!
//! ### Debug Mode Activation
//!
//! ```vb
//! Public DebugMode As Boolean
//!
//! Sub Main()
//!     Dim cmdLine As String
//!     cmdLine = LCase(Trim(Command()))
//!     
//!     DebugMode = (InStr(cmdLine, "/debug") > 0) Or _
//!                 (InStr(cmdLine, "-debug") > 0) Or _
//!                 (InStr(cmdLine, "/d") > 0)
//!     
//!     If DebugMode Then
//!         MsgBox "Debug mode enabled"
//!     End If
//! End Sub
//! ```
//!
//! ### Multiple File Processing
//!
//! ```vb
//! Sub Main()
//!     Dim files() As String
//!     Dim i As Integer
//!     
//!     files = GetCommandLineFiles()
//!     
//!     For i = LBound(files) To UBound(files)
//!         ProcessFile files(i)
//!     Next i
//! End Sub
//!
//! Function GetCommandLineFiles() As String()
//!     Dim cmdLine As String
//!     Dim args As Collection
//!     Dim result() As String
//!     Dim i As Integer
//!     Dim count As Integer
//!     
//!     Set args = ParseCommandLine()
//!     
//!     ' Count files (skip switches)
//!     For i = 1 To args.Count
//!         If Left(args(i), 1) <> "/" And Left(args(i), 1) <> "-" Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     If count > 0 Then
//!         ReDim result(1 To count)
//!         count = 0
//!         
//!         For i = 1 To args.Count
//!             If Left(args(i), 1) <> "/" And Left(args(i), 1) <> "-" Then
//!                 count = count + 1
//!                 result(count) = args(i)
//!             End If
//!         Next i
//!     End If
//!     
//!     GetCommandLineFiles = result
//! End Function
//! ```
//!
//! ### Automation Mode
//!
//! ```vb
//! Sub Main()
//!     If HasSwitch("auto") Or HasSwitch("batch") Then
//!         ' Run in automated mode without UI
//!         RunBatchProcess
//!         End
//!     Else
//!         ' Show normal UI
//!         Form1.Show
//!     End If
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Complex Argument Parser
//!
//! ```vb
//! Type CommandLineArg
//!     Name As String
//!     Value As String
//!     IsSwitch As Boolean
//! End Type
//!
//! Function ParseAdvancedCommandLine() As Collection
//!     Dim args As New Collection
//!     Dim cmdLine As String
//!     Dim tokens As Collection
//!     Dim i As Integer
//!     Dim token As String
//!     Dim arg As CommandLineArg
//!     
//!     cmdLine = Command()
//!     Set tokens = ParseCommandLine()
//!     
//!     For i = 1 To tokens.Count
//!         token = tokens(i)
//!         
//!         If Left(token, 1) = "/" Or Left(token, 1) = "-" Then
//!             arg.IsSwitch = True
//!             
//!             ' Remove leading / or -
//!             token = Mid(token, 2)
//!             
//!             ' Check for name:value format
//!             If InStr(token, ":") > 0 Then
//!                 arg.Name = Left(token, InStr(token, ":") - 1)
//!                 arg.Value = Mid(token, InStr(token, ":") + 1)
//!             ElseIf InStr(token, "=") > 0 Then
//!                 arg.Name = Left(token, InStr(token, "=") - 1)
//!                 arg.Value = Mid(token, InStr(token, "=") + 1)
//!             Else
//!                 arg.Name = token
//!                 arg.Value = "True"
//!             End If
//!         Else
//!             arg.IsSwitch = False
//!             arg.Name = ""
//!             arg.Value = token
//!         End If
//!         
//!         args.Add arg
//!     Next i
//!     
//!     Set ParseAdvancedCommandLine = args
//! End Function
//! ```
//!
//! ### Environment Variable Expansion
//!
//! ```vb
//! Function ExpandCommandLine() As String
//!     Dim cmdLine As String
//!     Dim startPos As Integer
//!     Dim endPos As Integer
//!     Dim varName As String
//!     Dim varValue As String
//!     
//!     cmdLine = Command()
//!     
//!     ' Expand %VARIABLE% syntax
//!     Do
//!         startPos = InStr(cmdLine, "%")
//!         If startPos = 0 Then Exit Do
//!         
//!         endPos = InStr(startPos + 1, cmdLine, "%")
//!         If endPos = 0 Then Exit Do
//!         
//!         varName = Mid(cmdLine, startPos + 1, endPos - startPos - 1)
//!         varValue = Environ(varName)
//!         
//!         cmdLine = Left(cmdLine, startPos - 1) & varValue & Mid(cmdLine, endPos + 1)
//!     Loop
//!     
//!     ExpandCommandLine = cmdLine
//! End Function
//! ```
//!
//! ### Help Text Display
//!
//! ```vb
//! Sub Main()
//!     If HasSwitch("?") Or HasSwitch("help") Then
//!         ShowHelp
//!         End
//!     End If
//!     
//!     ' Normal startup
//!     Form1.Show
//! End Sub
//!
//! Sub ShowHelp()
//!     Dim helpText As String
//!     
//!     helpText = "MyApp - Command Line Options" & vbCrLf & vbCrLf
//!     helpText = helpText & "/debug        Enable debug mode" & vbCrLf
//!     helpText = helpText & "/config:file  Load configuration from file" & vbCrLf
//!     helpText = helpText & "/silent       Run in silent mode" & vbCrLf
//!     helpText = helpText & "/auto         Run in automated mode" & vbCrLf
//!     helpText = helpText & "/help or /?   Show this help" & vbCrLf
//!     
//!     MsgBox helpText, vbInformation
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeGetCommand() As String
//!     On Error GoTo ErrorHandler
//!     
//!     SafeGetCommand = Command()
//!     Exit Function
//!     
//! ErrorHandler:
//!     ' Command() rarely fails, but handle just in case
//!     SafeGetCommand = ""
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Command()` is a fast function with minimal overhead
//! - Result is cached, so multiple calls don't re-query the OS
//! - Consider caching the result in a module-level variable if used frequently
//! - Parsing complex command lines can be expensive; cache parsed results
//!
//! ## Best Practices
//!
//! ### Cache the Result
//!
//! ```vb
//! Public g_CommandLine As String
//!
//! Sub Main()
//!     g_CommandLine = Command()
//!     
//!     ' Use g_CommandLine throughout the application
//!     If InStr(g_CommandLine, "/debug") > 0 Then
//!         ' ...
//!     End If
//! End Sub
//! ```
//!
//! ### Validate Arguments
//!
//! ```vb
//! Sub Main()
//!     Dim cmdLine As String
//!     cmdLine = Command()
//!     
//!     If cmdLine <> "" Then
//!         If Not ValidateCommandLine(cmdLine) Then
//!             MsgBox "Invalid command line arguments", vbCritical
//!             End
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ### Use `Sub Main()` for Command Line Apps
//!
//! ```vb
//! Sub Main()
//!     ' Process command line before showing any UI
//!     ProcessCommandLine
//!     
//!     ' Then show UI or continue processing
//!     Form1.Show
//! End Sub
//! ```
//!
//! ## Limitations
//!
//! - Returns only arguments, not the executable path (use App.Path and App.EXEName instead)
//! - No built-in parsing; returns raw string
//! - Quote handling is not automatic
//! - Limited to approximately 32KB of text on some Windows versions
//! - No standard format for arguments (application must define its own conventions)
//! - Different from C/C++ argv[] which provides separate argument array
//!
//! ## Related Functions and Properties
//!
//! - `App.Path`: Returns the path where the application is located
//! - `App.EXEName`: Returns the executable filename without extension
//! - `Environ`: Gets environment variable values
//! - `Shell`: Executes external programs with command lines
//!
//! ## Platform Considerations
//!
//! - Windows: Uses `GetCommandLine` API internally
//! - Command line length limits vary by Windows version
//! - Unicode characters may require special handling
//! - Some special characters may need escaping in batch files

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn command_basic() {
        let source = r"
args = Command()
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("args"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_in_assignment() {
        let source = r"
Dim cmdLine As String
cmdLine = Command()
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("cmdLine"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cmdLine"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_in_if_statement() {
        let source = r#"
If Command() <> "" Then
    ProcessArgs
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Command"),
                        LeftParenthesis,
                        ArgumentList,
                        RightParenthesis,
                    },
                    Whitespace,
                    InequalityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessArgs"),
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
    fn command_with_trim() {
        let source = r"
args = Trim(Command())
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("args"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Trim"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
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
    fn command_in_instr() {
        let source = r#"
If InStr(Command(), "/debug") > 0 Then
    DebugMode = True
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("InStr"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Command"),
                                    LeftParenthesis,
                                    ArgumentList,
                                    RightParenthesis,
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"/debug\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
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
                            Identifier ("DebugMode"),
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
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn command_in_sub_main() {
        let source = r"
Sub Main()
    Dim args As String
    args = Command()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        Identifier ("args"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("args"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Command"),
                            LeftParenthesis,
                            ArgumentList,
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
    fn command_in_msgbox() {
        let source = r#"
MsgBox "Args: " & Command()
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Args: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Command"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn command_with_lcase() {
        let source = r"
cmdLine = LCase(Command())
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cmdLine"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("LCase"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
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
    fn command_in_function() {
        let source = r"
Function GetArgs() As String
    GetArgs = Command()
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetArgs"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetArgs"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Command"),
                            LeftParenthesis,
                            ArgumentList,
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
    fn command_in_select_case() {
        let source = r#"
Select Case Command()
    Case "/help"
        ShowHelp
    Case "/debug"
        EnableDebug
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList,
                    RightParenthesis,
                },
                Newline,
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\"/help\""),
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
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\"/debug\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("EnableDebug"),
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
    fn command_empty_check() {
        let source = r#"
If Command() = "" Then
    MsgBox "No arguments"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Command"),
                        LeftParenthesis,
                        ArgumentList,
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
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
                        StringLiteral ("\"No arguments\""),
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
    fn command_with_split() {
        let source = r#"
args = Split(Command(), " ")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("args"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Split"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\" \""),
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
    fn command_in_do_loop() {
        let source = r#"
Do While Command() <> ""
    Process
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Command"),
                        LeftParenthesis,
                        ArgumentList,
                        RightParenthesis,
                    },
                    Whitespace,
                    InequalityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\""),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Process"),
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn command_with_left() {
        let source = r#"
If Left(Command(), 1) = "/" Then
    ProcessSwitch
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Left"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Command"),
                                    LeftParenthesis,
                                    ArgumentList,
                                    RightParenthesis,
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
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"/\""),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessSwitch"),
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
    fn command_in_replace() {
        let source = r#"
args = Replace(Command(), "/", "-")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("args"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Replace"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"/\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"-\""),
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
    fn command_len_check() {
        let source = r"
If Len(Command()) > 0 Then
    ParseArgs
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
                    CallExpression {
                        LenKeyword,
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Command"),
                                    LeftParenthesis,
                                    ArgumentList,
                                    RightParenthesis,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
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
                    CallStatement {
                        Identifier ("ParseArgs"),
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
    fn command_in_concatenation() {
        let source = r#"
fullCmd = App.EXEName & " " & Command()
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("fullCmd"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        MemberAccessExpression {
                            Identifier ("App"),
                            PeriodOperator,
                            Identifier ("EXEName"),
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\" \""),
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("Command"),
                        LeftParenthesis,
                        ArgumentList,
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_multiple_calls() {
        let source = r"
cmd1 = Command()
cmd2 = Command()
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cmd1"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList,
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cmd2"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_in_for_loop() {
        let source = r"
For i = 1 To Len(Command())
    ch = Mid(Command(), i, 1)
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
                CallExpression {
                    LenKeyword,
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ch"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            MidKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Command"),
                                        LeftParenthesis,
                                        ArgumentList,
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
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
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn command_with_ucase() {
        let source = r"
cmdUpper = UCase(Command())
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cmdUpper"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("UCase"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
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
    fn command_in_comparison() {
        let source = r#"
result = (Command() = "/auto")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
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
                ParenthesizedExpression {
                    LeftParenthesis,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("Command"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"/auto\""),
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_with_right() {
        let source = r#"
If Right(Command(), 4) = ".txt" Then
    ProcessTextFile
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Right"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Command"),
                                    LeftParenthesis,
                                    ArgumentList,
                                    RightParenthesis,
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("4"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\".txt\""),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessTextFile"),
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
    fn command_with_mid() {
        let source = r"
part = Mid(Command(), 2, 5)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("part"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    MidKeyword,
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Command"),
                                LeftParenthesis,
                                ArgumentList,
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("5"),
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
    fn command_with_whitespace() {
        let source = r"
args = Command( )
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("args"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Command"),
                    LeftParenthesis,
                    ArgumentList {
                        Whitespace,
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn command_in_print() {
        let source = r#"
Print "Command line: "; Command()
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PrintStatement {
                PrintKeyword,
                Whitespace,
                StringLiteral ("\"Command line: \""),
                Semicolon,
                Whitespace,
                Identifier ("Command"),
                LeftParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }
}
