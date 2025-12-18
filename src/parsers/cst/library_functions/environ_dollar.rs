//! # `Environ$` Function
//!
//! Returns the string value associated with an environment variable.
//!
//! ## Syntax
//!
//! ```vb6
//! Environ$(envstring)
//! Environ$(number)
//! ```
//!
//! ## Parameters
//!
//! - `envstring`: A string expression containing the name of an environment variable.
//! - `number`: A numeric expression corresponding to the numeric order of an environment string in the environment-string table. The number argument can be any numeric expression, but is rounded to a whole number before it is evaluated.
//!
//! ## Return Value
//!
//! Returns a `String` containing the text assigned to the specified environment variable. If the environment variable doesn't exist, returns an empty string.
//!
//! ## Remarks
//!
//! The `Environ$` function returns the string assigned to the specified environment variable from the operating system's environment-string table. This function cannot be used on the left side of an assignment statement.
//!
//! When using a numeric argument, `Environ$` returns the string that occupies that numeric position in the environment table. In this case, `Environ$` returns all the text including the equal sign (=). If there's no environment string at the specified position, `Environ$` returns a zero-length string.
//!
//! When using a string argument, if the environment variable doesn't exist, a zero-length string is returned.
//!
//! ## Typical Uses
//!
//! ### Example 1: Getting System Path
//! ```vb6
//! Dim systemPath As String
//! systemPath = Environ$("PATH")
//! ```
//!
//! ### Example 2: Getting Temp Directory
//! ```vb6
//! Dim tempDir As String
//! tempDir = Environ$("TEMP")
//! ```
//!
//! ### Example 3: Getting User Name
//! ```vb6
//! Dim userName As String
//! userName = Environ$("USERNAME")
//! ```
//!
//! ### Example 4: Iterating Environment Variables
//! ```vb6
//! Dim i As Integer
//! Dim envVar As String
//! i = 1
//! Do
//!     envVar = Environ$(i)
//!     If envVar <> "" Then Debug.Print envVar
//!     i = i + 1
//! Loop While envVar <> ""
//! ```
//!
//! ## Common Usage Patterns
//!
//! ### Getting Application Data Path
//! ```vb6
//! Dim appDataPath As String
//! appDataPath = Environ$("APPDATA")
//! If appDataPath <> "" Then
//!     appDataPath = appDataPath & "\MyApp\"
//! End If
//! ```
//!
//! ### Getting User Profile Directory
//! ```vb6
//! Dim userProfile As String
//! userProfile = Environ$("USERPROFILE")
//! configFile = userProfile & "\config.ini"
//! ```
//!
//! ### Checking for Development Environment
//! ```vb6
//! Dim devMode As Boolean
//! devMode = (Environ$("DEV_MODE") = "1")
//! If devMode Then
//!     Debug.Print "Running in development mode"
//! End If
//! ```
//!
//! ### Building Full Path with Temp Directory
//! ```vb6
//! Dim tempFile As String
//! tempFile = Environ$("TEMP") & "\tempdata.tmp"
//! Open tempFile For Output As #1
//! ```
//!
//! ### Getting System Drive
//! ```vb6
//! Dim systemDrive As String
//! systemDrive = Environ$("SystemDrive")
//! logPath = systemDrive & "\Logs\app.log"
//! ```
//!
//! ### Listing All Environment Variables
//! ```vb6
//! Dim idx As Integer
//! Dim envEntry As String
//! For idx = 1 To 255
//!     envEntry = Environ$(idx)
//!     If envEntry = "" Then Exit For
//!     List1.AddItem envEntry
//! Next idx
//! ```
//!
//! ### Cross-Platform Path Separator
//! ```vb6
//! Dim pathSep As String
//! If Environ$("OS") Like "Windows*" Then
//!     pathSep = "\"
//! Else
//!     pathSep = "/"
//! End If
//! ```
//!
//! ### Getting Computer Name
//! ```vb6
//! Dim computerName As String
//! computerName = Environ$("COMPUTERNAME")
//! If computerName = "" Then computerName = Environ$("HOSTNAME")
//! ```
//!
//! ### Building Log File Path with User Name
//! ```vb6
//! Dim logFile As String
//! logFile = "C:\Logs\" & Environ$("USERNAME") & ".log"
//! Open logFile For Append As #1
//! Print #1, Now & " - User logged in"
//! Close #1
//! ```
//!
//! ### Checking if Variable Exists
//! ```vb6
//! Dim dbServer As String
//! dbServer = Environ$("DB_SERVER")
//! If dbServer = "" Then
//!     dbServer = "localhost"  ' Default value
//! End If
//! ```
//!
//! ## Related Functions
//!
//! - `Environ`: Non-string variant (returns Variant)
//! - `Command$`: Gets command-line arguments
//! - `CurDir$`: Gets current directory
//! - `GetSetting`: Reads application settings from registry
//! - `Dir$`: Lists files in directory
//!
//! ## Best Practices
//!
//! 1. Always check if the returned value is empty before using it
//! 2. Use string argument form for better code readability
//! 3. Cache frequently accessed environment variables
//! 4. Be aware of case sensitivity on different platforms
//! 5. Avoid modifying environment variables from VB6 (use shell APIs instead)
//! 6. Use proper path combining (avoid double backslashes)
//! 7. Consider using `GetEnvironmentVariable` API for more control
//! 8. Remember that environment variables persist only for the process lifetime
//! 9. Use constants for commonly used environment variable names
//! 10. Validate paths returned from environment variables before using them
//!
//! ## Performance Considerations
//!
//! - Environment variable lookup is relatively fast
//! - Iterating all variables with numeric index is slower than direct lookup
//! - Consider caching values if used frequently in loops
//! - No significant performance difference between `Environ$` and `Environ`
//!
//! ## Platform Differences
//!
//! | Platform | Notes |
//! |----------|-------|
//! | Windows 95/98 | Limited environment space (may fail with many variables) |
//! | Windows NT/2000/XP | Larger environment space, more reliable |
//! | Windows Vista+ | User and system environment variables separated |
//! | Wine/Linux | May return different variables, case sensitivity differs |
//!
//! ## Common Environment Variables
//!
//! | Variable | Description |
//! |----------|-------------|
//! | `PATH` | System search path for executables |
//! | `TEMP` or `TMP` | Temporary files directory |
//! | `APPDATA` | Application data folder (Windows) |
//! | `USERPROFILE` | User's home directory (Windows) |
//! | `USERNAME` | Current user's login name |
//! | `COMPUTERNAME` | Computer's network name |
//! | `SystemDrive` | Drive letter of system installation |
//! | `SystemRoot` | Windows installation directory |
//! | `HOMEDRIVE` | User's home drive letter |
//! | `HOMEPATH` | User's home directory path |
//!
//! ## Common Pitfalls
//!
//! - Not checking for empty string return values
//! - Assuming environment variable names are case-insensitive on all platforms
//! - Using numeric index without checking for empty string to detect end
//! - Creating paths with double backslashes when concatenating
//! - Assuming all common variables exist on all systems
//! - Not handling missing required environment variables gracefully
//!
//! ## Limitations
//!
//! - Cannot be used to set environment variables (use Windows API)
//! - Environment changes don't persist beyond process lifetime
//! - Limited to current process's environment space
//! - Some variables may be protected or unavailable depending on permissions
//! - Variable availability differs between operating systems

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn environ_dollar_simple() {
        let source = r#"
Sub Main()
    path = Environ$("PATH")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim tempDir As String
    tempDir = Environ$("TEMP")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_concatenation() {
        let source = r#"
Sub Main()
    configPath = Environ$("APPDATA") & "\MyApp\config.ini"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Environ$("DEV_MODE") = "1" Then
        Debug.Print "Development mode"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_numeric_index() {
        let source = r#"
Sub Main()
    Dim i As Integer
    For i = 1 To 100
        envVar = Environ$(i)
        If envVar = "" Then Exit For
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_user_profile() {
        let source = r#"
Sub Main()
    userDir = Environ$("USERPROFILE")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_temp_file() {
        let source = r#"
Sub Main()
    tempFile = Environ$("TEMP") & "\data.tmp"
    Open tempFile For Output As #1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_username() {
        let source = r#"
Sub Main()
    currentUser = Environ$("USERNAME")
    logFile = "C:\Logs\" & currentUser & ".log"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_computer_name() {
        let source = r#"
Sub Main()
    machine = Environ$("COMPUTERNAME")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_default_value() {
        let source = r#"
Sub Main()
    dbServer = Environ$("DB_SERVER")
    If dbServer = "" Then dbServer = "localhost"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_system_drive() {
        let source = r#"
Sub Main()
    sysDrive = Environ$("SystemDrive")
    logPath = sysDrive & "\Logs\app.log"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_with_empty_check() {
        let source = r#"
Sub Main()
    appData = Environ$("APPDATA")
    If appData <> "" Then
        appData = appData & "\MyApp\"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_list_all() {
        let source = r"
Sub Main()
    Dim idx As Integer
    Dim entry As String
    idx = 1
    entry = Environ$(idx)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_multiple_uses() {
        let source = r#"
Sub Main()
    user = Environ$("USERNAME")
    comp = Environ$("COMPUTERNAME")
    msg = user & "@" & comp
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_in_function() {
        let source = r#"
Function GetTempPath() As String
    GetTempPath = Environ$("TEMP")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_with_format() {
        let source = r#"
Sub Main()
    result = "User: " & Environ$("USERNAME")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_select_case() {
        let source = r#"
Sub Main()
    osType = Environ$("OS")
    Select Case osType
        Case "Windows_NT"
            Debug.Print "NT-based"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_in_loop() {
        let source = r"
Sub Main()
    Dim i As Integer
    For i = 1 To 50
        v = Environ$(i)
    Next i
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_with_len() {
        let source = r#"
Sub Main()
    pathVar = Environ$("PATH")
    pathLen = Len(pathVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }

    #[test]
    fn environ_dollar_path_building() {
        let source = r#"
Sub Main()
    userPath = Environ$("USERPROFILE") & "\Documents\data.txt"
    Open userPath For Input As #1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Environ$"));
    }
}
