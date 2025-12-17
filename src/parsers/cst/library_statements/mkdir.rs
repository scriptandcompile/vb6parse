//! # `MkDir` Statement
//!
//! Creates a new directory or folder.
//!
//! ## Syntax
//!
//! ```vb
//! MkDir path
//! ```
//!
//! - `path`: Required. String expression that identifies the directory or folder to be created. May include drive.
//!   If no drive is specified, `MkDir` creates the new directory or folder on the current drive.
//!
//! ## Remarks
//!
//! - An error occurs if you try to create a directory or folder that already exists
//! - The `path` argument can include absolute or relative paths
//! - You can use `MkDir` to create nested directories by creating parent directories first
//! - On Windows systems, both forward slashes (/) and backslashes (\) can be used as path separators
//! - The directory name can include the drive letter
//! - UNC paths are supported on network drives
//!
//! ## Examples
//!
//! ```vb
//! ' Create a directory in the current directory
//! MkDir "MyNewFolder"
//!
//! ' Create a directory with full path
//! MkDir "C:\Program Files\MyApp"
//!
//! ' Create a directory on another drive
//! MkDir "D:\Data\Reports"
//!
//! ' Create nested directories (parent must exist first)
//! MkDir "C:\Temp"
//! MkDir "C:\Temp\Logs"
//! MkDir "C:\Temp\Logs\Archive"
//!
//! ' Create directory on network drive
//! MkDir "\\Server\Share\NewFolder"
//! ```
//!
//! ## Reference
//!
//! [MkDir Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mkdir-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `MkDir` statement.
    pub(crate) fn parse_mkdir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::MkDirStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // MkDir statement tests
    #[test]
    fn mkdir_simple() {
        let source = r#"
Sub Test()
    MkDir "NewFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("MkDirKeyword"));
        assert!(debug.contains("NewFolder"));
    }

    #[test]
    fn mkdir_at_module_level() {
        let source = r#"MkDir "C:\Temp""#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn mkdir_with_full_path() {
        let source = r#"
Sub Test()
    MkDir "C:\Program Files\MyApp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("Program Files"));
    }

    #[test]
    fn mkdir_with_variable() {
        let source = r#"
Sub Test()
    MkDir folderPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("folderPath"));
    }

    #[test]
    fn mkdir_preserves_whitespace() {
        let source = "    MkDir    \"MyFolder\"    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    MkDir    \"MyFolder\"    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn mkdir_with_comment() {
        let source = r#"
Sub Test()
    MkDir "Logs" ' Create logs directory
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn mkdir_in_if_statement() {
        let source = r#"
Sub Test()
    If Not dirExists Then
        MkDir "C:\Data"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn mkdir_inline_if() {
        let source = r#"
Sub Test()
    If needsDir Then MkDir "Output"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn multiple_mkdir_statements() {
        let source = r#"
Sub CreateDirs()
    MkDir "C:\Temp"
    MkDir "C:\Temp\Logs"
    MkDir "C:\Temp\Data"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("MkDirStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn mkdir_with_concatenation() {
        let source = r#"
Sub Test()
    MkDir basePath & "\Subfolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("basePath"));
    }

    #[test]
    fn mkdir_unc_path() {
        let source = r#"
Sub Test()
    MkDir "\\Server\Share\NewFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("Server"));
    }

    #[test]
    fn mkdir_with_function_call() {
        let source = r#"
Sub Test()
    MkDir App.Path & "\Data"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
        assert!(debug.contains("App"));
        assert!(debug.contains("Path"));
    }
}
