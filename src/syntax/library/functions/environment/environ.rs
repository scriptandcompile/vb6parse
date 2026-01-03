//! # Environ Function
//!
//! Returns the String value associated with an operating system environment variable.
//!
//! ## Syntax
//!
//! ```vb
//! Environ(envstring | number)
//! ```
//!
//! ## Parameters
//!
//! - **envstring**: Optional (if number provided). String expression containing the name of an
//!   environment variable.
//! - **number**: Optional (if envstring provided). Numeric expression corresponding to the
//!   numeric order of an environment string in the environment-string table. The number argument
//!   can be any numeric expression, but is rounded to a whole number before it is evaluated.
//!
//! ## Return Value
//!
//! Returns a String containing the value assigned to envstring or the environment variable at
//! position number. Returns a zero-length string ("") if envstring is not found or if there is
//! no environment string at position number.
//!
//! ## Remarks
//!
//! The `Environ` function retrieves values from the operating system's environment variables.
//! Environment variables are system-level or user-level settings that provide information about
//! the operating system environment.
//!
//! **Important Characteristics:**
//!
//! - Reads environment variables from the operating system
//! - Can access by name (string) or position (number)
//! - Case-insensitive on Windows
//! - Returns empty string if variable not found
//! - Position-based access starts at 1 (not 0)
//! - Number of environment variables varies by system
//! - Environment changes during execution are not reflected
//! - Snapshot taken at application start
//! - Cannot modify environment variables (read-only)
//! - Different users may have different environment variables
//!
//! ## Common Environment Variables
//!
//! **Windows:**
//! - `PATH` - Executable search path
//! - `TEMP` or `TMP` - Temporary files directory
//! - `USERNAME` - Current user name
//! - `USERPROFILE` - User's profile directory
//! - `COMPUTERNAME` - Computer name
//! - `SYSTEMROOT` - Windows system directory
//! - `PROGRAMFILES` - Program Files directory
//! - `HOMEDRIVE` - User's home drive (e.g., C:)
//! - `HOMEPATH` - User's home directory path
//! - `APPDATA` - Application data directory
//! - `WINDIR` - Windows directory
//! - `PROCESSOR_ARCHITECTURE` - CPU architecture
//! - `NUMBER_OF_PROCESSORS` - Number of CPU cores
//! - `OS` - Operating system name
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get environment variable by name
//! Dim userName As String
//! userName = Environ("USERNAME")
//! MsgBox "Current user: " & userName
//!
//! ' Get temp directory
//! Dim tempDir As String
//! tempDir = Environ("TEMP")
//!
//! ' Get by position
//! Dim firstEnvVar As String
//! firstEnvVar = Environ(1)
//! ```
//!
//! ### Check if Variable Exists
//!
//! ```vb
//! Function EnvironVarExists(varName As String) As Boolean
//!     EnvironVarExists = (Len(Environ(varName)) > 0)
//! End Function
//!
//! ' Usage
//! If EnvironVarExists("JAVA_HOME") Then
//!     MsgBox "Java is configured"
//! Else
//!     MsgBox "Java not found"
//! End If
//! ```
//!
//! ### Build File Paths
//!
//! ```vb
//! Function GetTempFilePath(fileName As String) As String
//!     Dim tempDir As String
//!     tempDir = Environ("TEMP")
//!     
//!     If Right(tempDir, 1) <> "\" Then
//!         tempDir = tempDir & "\"
//!     End If
//!     
//!     GetTempFilePath = tempDir & fileName
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Get User Directories
//!
//! ```vb
//! Function GetUserProfile() As String
//!     GetUserProfile = Environ("USERPROFILE")
//! End Function
//!
//! Function GetAppDataPath() As String
//!     GetAppDataPath = Environ("APPDATA")
//! End Function
//!
//! Function GetDesktopPath() As String
//!     GetDesktopPath = Environ("USERPROFILE") & "\Desktop"
//! End Function
//!
//! Function GetMyDocuments() As String
//!     GetMyDocuments = Environ("USERPROFILE") & "\Documents"
//! End Function
//! ```
//!
//! ### System Information
//!
//! ```vb
//! Function GetComputerName() As String
//!     GetComputerName = Environ("COMPUTERNAME")
//! End Function
//!
//! Function GetProcessorCount() As Integer
//!     Dim procCount As String
//!     procCount = Environ("NUMBER_OF_PROCESSORS")
//!     
//!     If IsNumeric(procCount) Then
//!         GetProcessorCount = CInt(procCount)
//!     Else
//!         GetProcessorCount = 1
//!     End If
//! End Function
//!
//! Function GetSystemArchitecture() As String
//!     GetSystemArchitecture = Environ("PROCESSOR_ARCHITECTURE")
//! End Function
//! ```
//!
//! ### List All Environment Variables
//!
//! ```vb
//! Sub ListAllEnvironmentVariables()
//!     Dim i As Integer
//!     Dim envVar As String
//!     
//!     i = 1
//!     Do
//!         envVar = Environ(i)
//!         If envVar = "" Then Exit Do
//!         
//!         Debug.Print i & ": " & envVar
//!         i = i + 1
//!     Loop
//! End Sub
//! ```
//!
//! ### Parse Environment Variable
//!
//! ```vb
//! Function GetEnvironVarName(envString As String) As String
//!     Dim equalPos As Integer
//!     equalPos = InStr(envString, "=")
//!     
//!     If equalPos > 0 Then
//!         GetEnvironVarName = Left(envString, equalPos - 1)
//!     Else
//!         GetEnvironVarName = ""
//!     End If
//! End Function
//!
//! Function GetEnvironVarValue(envString As String) As String
//!     Dim equalPos As Integer
//!     equalPos = InStr(envString, "=")
//!     
//!     If equalPos > 0 Then
//!         GetEnvironVarValue = Mid(envString, equalPos + 1)
//!     Else
//!         GetEnvironVarValue = ""
//!     End If
//! End Function
//! ```
//!
//! ### Safe Path Construction
//!
//! ```vb
//! Function BuildSafePath(envVar As String, subPath As String) As String
//!     Dim basePath As String
//!     basePath = Environ(envVar)
//!     
//!     If basePath = "" Then
//!         BuildSafePath = ""
//!         Exit Function
//!     End If
//!     
//!     ' Ensure path ends with backslash
//!     If Right(basePath, 1) <> "\" Then
//!         basePath = basePath & "\"
//!     End If
//!     
//!     ' Remove leading backslash from subPath if present
//!     If Left(subPath, 1) = "\" Then
//!         subPath = Mid(subPath, 2)
//!     End If
//!     
//!     BuildSafePath = basePath & subPath
//! End Function
//! ```
//!
//! ### Configuration File Paths
//!
//! ```vb
//! Function GetConfigFilePath(appName As String, fileName As String) As String
//!     Dim appDataPath As String
//!     Dim configDir As String
//!     
//!     appDataPath = Environ("APPDATA")
//!     configDir = appDataPath & "\" & appName
//!     
//!     ' Create directory if it doesn't exist
//!     If Dir(configDir, vbDirectory) = "" Then
//!         MkDir configDir
//!     End If
//!     
//!     GetConfigFilePath = configDir & "\" & fileName
//! End Function
//! ```
//!
//! ### Search PATH Variable
//!
//! ```vb
//! Function FindInPath(executable As String) As String
//!     Dim pathVar As String
//!     Dim paths() As String
//!     Dim i As Integer
//!     Dim testPath As String
//!     
//!     pathVar = Environ("PATH")
//!     paths = Split(pathVar, ";")
//!     
//!     For i = LBound(paths) To UBound(paths)
//!         testPath = paths(i) & "\" & executable
//!         If Dir(testPath) <> "" Then
//!             FindInPath = testPath
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     FindInPath = ""
//! End Function
//! ```
//!
//! ### Check Operating System
//!
//! ```vb
//! Function IsWindows() As Boolean
//!     Dim osVar As String
//!     osVar = UCase(Environ("OS"))
//!     IsWindows = (InStr(osVar, "WINDOWS") > 0)
//! End Function
//!
//! Function GetWindowsDirectory() As String
//!     GetWindowsDirectory = Environ("WINDIR")
//! End Function
//!
//! Function GetSystemRoot() As String
//!     GetSystemRoot = Environ("SYSTEMROOT")
//! End Function
//! ```
//!
//! ### Program Files Paths
//!
//! ```vb
//! Function GetProgramFilesPath() As String
//!     GetProgramFilesPath = Environ("PROGRAMFILES")
//! End Function
//!
//! Function GetProgramFilesX86Path() As String
//!     GetProgramFilesX86Path = Environ("PROGRAMFILES(X86)")
//! End Function
//!
//! Function FindProgramPath(programName As String) As String
//!     Dim progFiles As String
//!     Dim testPath As String
//!     
//!     ' Check Program Files
//!     progFiles = Environ("PROGRAMFILES")
//!     testPath = progFiles & "\" & programName
//!     If Dir(testPath, vbDirectory) <> "" Then
//!         FindProgramPath = testPath
//!         Exit Function
//!     End If
//!     
//!     ' Check Program Files (x86)
//!     progFiles = Environ("PROGRAMFILES(X86)")
//!     If progFiles <> "" Then
//!         testPath = progFiles & "\" & programName
//!         If Dir(testPath, vbDirectory) <> "" Then
//!             FindProgramPath = testPath
//!             Exit Function
//!         End If
//!     End If
//!     
//!     FindProgramPath = ""
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Environment Variable Dictionary
//!
//! ```vb
//! Function GetEnvironmentDictionary() As Collection
//!     Dim dict As New Collection
//!     Dim i As Integer
//!     Dim envVar As String
//!     Dim varName As String
//!     Dim varValue As String
//!     Dim equalPos As Integer
//!     
//!     i = 1
//!     Do
//!         envVar = Environ(i)
//!         If envVar = "" Then Exit Do
//!         
//!         equalPos = InStr(envVar, "=")
//!         If equalPos > 0 Then
//!             varName = Left(envVar, equalPos - 1)
//!             varValue = Mid(envVar, equalPos + 1)
//!             
//!             On Error Resume Next
//!             dict.Add varValue, UCase(varName)
//!             On Error GoTo 0
//!         End If
//!         
//!         i = i + 1
//!     Loop
//!     
//!     Set GetEnvironmentDictionary = dict
//! End Function
//! ```
//!
//! ### Expand Environment Variables in String
//!
//! ```vb
//! Function ExpandEnvironmentString(inputString As String) As String
//!     Dim result As String
//!     Dim startPos As Integer
//!     Dim endPos As Integer
//!     Dim varName As String
//!     Dim varValue As String
//!     
//!     result = inputString
//!     
//!     ' Find %VAR% patterns
//!     Do
//!         startPos = InStr(result, "%")
//!         If startPos = 0 Then Exit Do
//!         
//!         endPos = InStr(startPos + 1, result, "%")
//!         If endPos = 0 Then Exit Do
//!         
//!         varName = Mid(result, startPos + 1, endPos - startPos - 1)
//!         varValue = Environ(varName)
//!         
//!         result = Left(result, startPos - 1) & varValue & Mid(result, endPos + 1)
//!     Loop
//!     
//!     ExpandEnvironmentString = result
//! End Function
//!
//! ' Usage
//! expandedPath = ExpandEnvironmentString("%TEMP%\myfile.txt")
//! ```
//!
//! ### Create Application Log File
//!
//! ```vb
//! Function CreateLogFile(appName As String) As String
//!     Dim logDir As String
//!     Dim logFile As String
//!     Dim dateStamp As String
//!     
//!     logDir = Environ("TEMP") & "\Logs"
//!     
//!     ' Create logs directory
//!     If Dir(logDir, vbDirectory) = "" Then
//!         MkDir logDir
//!     End If
//!     
//!     dateStamp = Format(Date, "yyyy-mm-dd")
//!     logFile = logDir & "\" & appName & "_" & dateStamp & ".log"
//!     
//!     CreateLogFile = logFile
//! End Function
//! ```
//!
//! ### Check Development Environment
//!
//! ```vb
//! Function IsDevelopmentEnvironment() As Boolean
//!     ' Check for common development environment variables
//!     IsDevelopmentEnvironment = (Len(Environ("VSCODE_PID")) > 0) Or _
//!                               (Len(Environ("TERM_PROGRAM")) > 0) Or _
//!                               (Len(Environ("VSAPPIDDIR")) > 0)
//! End Function
//!
//! Function GetJavaHome() As String
//!     GetJavaHome = Environ("JAVA_HOME")
//! End Function
//!
//! Function GetPythonPath() As String
//!     GetPythonPath = Environ("PYTHONPATH")
//! End Function
//! ```
//!
//! ### Build Connection String
//!
//! ```vb
//! Function BuildConnectionString() As String
//!     Dim server As String
//!     Dim database As String
//!     
//!     server = Environ("DB_SERVER")
//!     database = Environ("DB_NAME")
//!     
//!     If server = "" Then server = "localhost"
//!     If database = "" Then database = "default"
//!     
//!     BuildConnectionString = "Server=" & server & ";Database=" & database
//! End Function
//! ```
//!
//! ### Export Environment to File
//!
//! ```vb
//! Sub ExportEnvironmentToFile(filePath As String)
//!     Dim fileNum As Integer
//!     Dim i As Integer
//!     Dim envVar As String
//!     
//!     fileNum = FreeFile
//!     Open filePath For Output As #fileNum
//!     
//!     Print #fileNum, "Environment Variables"
//!     Print #fileNum, "Generated: " & Now
//!     Print #fileNum, String(80, "-")
//!     
//!     i = 1
//!     Do
//!         envVar = Environ(i)
//!         If envVar = "" Then Exit Do
//!         
//!         Print #fileNum, envVar
//!         i = i + 1
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Portable Path Builder
//!
//! ```vb
//! Function GetPortableAppPath(relativePath As String) As String
//!     Dim basePath As String
//!     
//!     ' Try to get from environment first
//!     basePath = Environ("APP_BASE_PATH")
//!     
//!     ' Fall back to current directory
//!     If basePath = "" Then
//!         basePath = App.Path
//!     End If
//!     
//!     If Right(basePath, 1) <> "\" Then
//!         basePath = basePath & "\"
//!     End If
//!     
//!     GetPortableAppPath = basePath & relativePath
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeEnviron(varName As String, Optional defaultValue As String = "") As String
//!     On Error Resume Next
//!     SafeEnviron = Environ(varName)
//!     
//!     If Err.Number <> 0 Or SafeEnviron = "" Then
//!         SafeEnviron = defaultValue
//!     End If
//! End Function
//!
//! Function GetEnvironWithFallback(preferredVar As String, fallbackVar As String) As String
//!     GetEnvironWithFallback = Environ(preferredVar)
//!     
//!     If GetEnvironWithFallback = "" Then
//!         GetEnvironWithFallback = Environ(fallbackVar)
//!     End If
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call): Invalid number argument (< 1)
//! - No error for missing variables (returns empty string)
//!
//! ## Performance Considerations
//!
//! - `Environ` is relatively fast (direct OS call)
//! - Environment snapshot taken at application start
//! - Cache frequently used values to avoid repeated calls
//! - Position-based enumeration stops at first empty string
//! - String comparison is case-insensitive on Windows
//!
//! ## Best Practices
//!
//! ### Always Check for Empty String
//!
//! ```vb
//! ' Good - Check before using
//! tempDir = Environ("TEMP")
//! If tempDir = "" Then
//!     tempDir = "C:\Temp"  ' Fallback
//! End If
//!
//! ' Avoid - Assuming variable exists
//! tempDir = Environ("TEMP")  ' May be empty!
//! ```
//!
//! ### Use Constants for Variable Names
//!
//! ```vb
//! ' Good - Constants for maintainability
//! Const ENV_TEMP = "TEMP"
//! Const ENV_USERNAME = "USERNAME"
//!
//! tempDir = Environ(ENV_TEMP)
//! userName = Environ(ENV_USERNAME)
//! ```
//!
//! ### Provide Defaults
//!
//! ```vb
//! Function GetTempDir() As String
//!     GetTempDir = Environ("TEMP")
//!     If GetTempDir = "" Then GetTempDir = Environ("TMP")
//!     If GetTempDir = "" Then GetTempDir = "C:\Temp"
//! End Function
//! ```
//!
//! ### Case Insensitive on Windows
//!
//! ```vb
//! ' All equivalent on Windows
//! userName = Environ("USERNAME")
//! userName = Environ("username")
//! userName = Environ("UserName")
//! ```
//!
//! ## Comparison with Other Methods
//!
//! ### Environ vs Registry
//!
//! ```vb
//! ' Environ - Quick, read-only access to environment
//! tempDir = Environ("TEMP")
//!
//! ' Registry - More control, can read/write, more complex
//! ' (Requires Windows API or Registry object)
//! ```
//!
//! ### Environ vs Command Line
//!
//! ```vb
//! ' Environ - Environment variables
//! userName = Environ("USERNAME")
//!
//! ' Command - Command line arguments
//! args = Command()
//! ```
//!
//! ## Limitations
//!
//! - Read-only access (cannot modify environment variables)
//! - Snapshot at application start (changes not reflected)
//! - Position-based enumeration order not guaranteed
//! - Cannot create or delete environment variables
//! - Limited to process environment (not system-wide)
//! - No wildcard or pattern matching
//! - Case-sensitive on Unix/Linux (VB6 primarily Windows)
//!
//! ## Related Functions
//!
//! - `Command`: Returns command-line arguments
//! - `CurDir`: Returns current directory
//! - `ChDir`: Changes current directory
//! - `App.Path`: Returns application path
//! - `Shell`: Executes external programs (can set environment)
//! - `GetSetting`: Reads application settings from registry
//! - `SaveSetting`: Writes application settings to registry

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn environ_basic_string() {
        let source = r#"
userName = Environ("USERNAME")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("userName"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"USERNAME\""),
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
    fn environ_with_number() {
        let source = r"
firstVar = Environ(1)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("firstVar"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
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
        ]);
    }

    #[test]
    fn environ_temp_dir() {
        let source = r#"
tempDir = Environ("TEMP")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("tempDir"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"TEMP\""),
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
    fn environ_in_function() {
        let source = r#"
Function GetUserProfile() As String
    GetUserProfile = Environ("USERPROFILE")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetUserProfile"),
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
                            Identifier ("GetUserProfile"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Environ"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"USERPROFILE\""),
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
    fn environ_existence_check() {
        let source = r#"
If Len(Environ("JAVA_HOME")) > 0 Then
    MsgBox "Java configured"
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
                        LenKeyword,
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Environ"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            StringLiteralExpression {
                                                StringLiteral ("\"JAVA_HOME\""),
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
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Java configured\""),
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
    fn environ_path_construction() {
        let source = r#"
appDataPath = Environ("APPDATA") & "\MyApp"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("appDataPath"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Environ"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"APPDATA\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\\MyApp\""),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn environ_loop_enumeration() {
        let source = r#"
i = 1
Do
    envVar = Environ(i)
    If envVar = "" Then Exit Do
    Debug.Print envVar
    i = i + 1
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("i"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("1"),
                },
                Newline,
            },
            DoStatement {
                DoKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("envVar"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Environ"),
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
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("envVar"),
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
                        Whitespace,
                        ExitStatement {
                            ExitKeyword,
                            Whitespace,
                            DoKeyword,
                            Newline,
                        },
                        Whitespace,
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintStatement {
                            PrintKeyword,
                            Whitespace,
                            Identifier ("envVar"),
                            Newline,
                        },
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                            },
                            Newline,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn environ_with_variable() {
        let source = r#"
varName = "PATH"
pathValue = Environ(varName)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("varName"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteralExpression {
                    StringLiteral ("\"PATH\""),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("pathValue"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("varName"),
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
    fn environ_comparison() {
        let source = r#"
If Environ("OS") = "Windows_NT" Then
    MsgBox "Windows NT"
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
                        Identifier ("Environ"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"OS\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"Windows_NT\""),
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
                        StringLiteral ("\"Windows NT\""),
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
    fn environ_msgbox() {
        let source = r#"
MsgBox "Computer: " & Environ("COMPUTERNAME")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Computer: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Environ"),
                LeftParenthesis,
                StringLiteral ("\"COMPUTERNAME\""),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn environ_select_case() {
        let source = r#"
Select Case UCase(Environ("OS"))
    Case "WINDOWS_NT"
        MsgBox "Windows NT"
    Case Else
        MsgBox "Other"
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
                    Identifier ("UCase"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Environ"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"OS\""),
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
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    StringLiteral ("\"WINDOWS_NT\""),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("MsgBox"),
                            Whitespace,
                            StringLiteral ("\"Windows NT\""),
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
                        CallStatement {
                            Identifier ("MsgBox"),
                            Whitespace,
                            StringLiteral ("\"Other\""),
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
    fn environ_error_handling() {
        let source = r#"
On Error Resume Next
value = Environ("CUSTOM_VAR")
If value = "" Then
    value = "default"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                ResumeKeyword,
                Whitespace,
                NextKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("value"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"CUSTOM_VAR\""),
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
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("value"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"default\""),
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
    fn environ_multiple_calls() {
        let source = r#"
user = Environ("USERNAME")
comp = Environ("COMPUTERNAME")
temp = Environ("TEMP")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("user"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"USERNAME\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("comp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"COMPUTERNAME\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"TEMP\""),
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
    fn environ_in_split() {
        let source = r#"
paths = Split(Environ("PATH"), ";")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("paths"),
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
                                Identifier ("Environ"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"PATH\""),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\";\""),
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
    fn environ_file_path() {
        let source = r#"
logFile = Environ("TEMP") & "\app.log"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("logFile"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Environ"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"TEMP\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\\app.log\""),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn environ_with_right() {
        let source = r#"
If Right(Environ("TEMP"), 1) <> "\" Then
    tempDir = Environ("TEMP") & "\"
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
                                    Identifier ("Environ"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            StringLiteralExpression {
                                                StringLiteral ("\"TEMP\""),
                                            },
                                        },
                                    },
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
                    InequalityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\\\" Then"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("tempDir"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Environ"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"TEMP\""),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\\\""),
                            },
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
    fn environ_isnumeric() {
        let source = r#"
procCount = Environ("NUMBER_OF_PROCESSORS")
If IsNumeric(procCount) Then
    cores = CInt(procCount)
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("procCount"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"NUMBER_OF_PROCESSORS\""),
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
                    Identifier ("IsNumeric"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("procCount"),
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
                            Identifier ("cores"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("CInt"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("procCount"),
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
    fn environ_for_loop() {
        let source = r#"
For i = 1 To 100
    envVar = Environ(i)
    If envVar = "" Then Exit For
    ProcessVar envVar
Next i
"#;
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("envVar"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Environ"),
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
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("envVar"),
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
                        Whitespace,
                        ExitStatement {
                            ExitKeyword,
                            Whitespace,
                            ForKeyword,
                            Newline,
                        },
                        Whitespace,
                        Identifier ("ProcessVar"),
                        Whitespace,
                        Identifier ("envVar"),
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
    fn environ_instr() {
        let source = r#"
If InStr(Environ("PATH"), "Java") > 0 Then
    MsgBox "Java in PATH"
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
                                    Identifier ("Environ"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            StringLiteralExpression {
                                                StringLiteral ("\"PATH\""),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"Java\""),
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
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Java in PATH\""),
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
    fn environ_with_dir() {
        let source = r#"
configDir = Environ("APPDATA") & "\MyApp"
If Dir(configDir, vbDirectory) = "" Then
    MkDir configDir
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("configDir"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Environ"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"APPDATA\""),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"\\MyApp\""),
                    },
                },
                Newline,
            },
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Dir"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("configDir"),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    Identifier ("vbDirectory"),
                                },
                            },
                        },
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
                    MkDirStatement {
                        Whitespace,
                        MkDirKeyword,
                        Whitespace,
                        Identifier ("configDir"),
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
    fn environ_default_value() {
        let source = r#"
value = Environ("CUSTOM_VAR")
If value = "" Then value = "default_value"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("value"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Environ"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"CUSTOM_VAR\""),
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
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("value"),
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
                Whitespace,
                AssignmentStatement {
                    IdentifierExpression {
                        Identifier ("value"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"default_value\""),
                    },
                    Newline,
                },
            },
        ]);
    }

    #[test]
    fn environ_debug_print() {
        let source = r#"
Debug.Print "User: " & Environ("USERNAME")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("Debug"),
                PeriodOperator,
                PrintKeyword,
                Whitespace,
                StringLiteral ("\"User: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Environ"),
                LeftParenthesis,
                StringLiteral ("\"USERNAME\""),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn environ_file_output() {
        let source = r#"
Open Environ("TEMP") & "\output.txt" For Output As #1
Print #1, Environ("USERNAME")
Close #1
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OpenStatement {
                OpenKeyword,
                Whitespace,
                Identifier ("Environ"),
                LeftParenthesis,
                StringLiteral ("\"TEMP\""),
                RightParenthesis,
                Whitespace,
                Ampersand,
                Whitespace,
                StringLiteral ("\"\\output.txt\""),
                Whitespace,
                ForKeyword,
                Whitespace,
                OutputKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Newline,
            },
            PrintStatement {
                PrintKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                Identifier ("Environ"),
                LeftParenthesis,
                StringLiteral ("\"USERNAME\""),
                RightParenthesis,
                Newline,
            },
            CloseStatement {
                CloseKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Newline,
            },
        ]);
    }

    #[test]
    fn environ_ucase() {
        let source = r#"
osName = UCase(Environ("OS"))
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("osName"),
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
                                Identifier ("Environ"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"OS\""),
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
        ]);
    }

    #[test]
    fn environ_collection() {
        let source = r#"
envDict.Add Environ("USERNAME"), "User"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("envDict"),
                PeriodOperator,
                Identifier ("Add"),
                Whitespace,
                Identifier ("Environ"),
                LeftParenthesis,
                StringLiteral ("\"USERNAME\""),
                RightParenthesis,
                Comma,
                Whitespace,
                StringLiteral ("\"User\""),
                Newline,
            },
        ]);
    }
}
