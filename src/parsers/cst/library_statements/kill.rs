use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    // VB6 Kill statement syntax:
    // - Kill pathname
    //
    // Deletes files from a disk.
    //
    // The Kill statement syntax has this part:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | pathname      | Required. String expression that specifies one or more file names to be deleted. May include directory or folder, and drive. |
    //
    // Remarks:
    // - Kill supports the use of multiple-character (*) and single-character (?) wildcards to specify multiple files.
    // - An error occurs if you try to use Kill to delete an open file.
    // - To remove a directory or folder, use the RmDir statement.
    //
    // Examples:
    // ```vb
    // Kill "C:\DATA.TXT"
    // Kill "C:\*.TXT"           ' Delete all .txt files
    // Kill "C:\TEST?.TXT"       ' Delete TEST1.TXT, TESTA.TXT, etc.
    // Kill App.Path & "\temp.dat"
    // Kill myFileName
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/kill-statement)
    pub(crate) fn parse_kill_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::KillStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn kill_simple() {
        let source = r#"
Sub Test()
    Kill "C:\DATA.TXT"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("DATA.TXT"));
    }

    #[test]
    fn kill_module_level() {
        let source = "Kill \"temp.dat\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("temp.dat"));
    }

    #[test]
    fn kill_with_wildcard() {
        let source = r#"
Sub Test()
    Kill "C:\*.TXT"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("*.TXT"));
    }

    #[test]
    fn kill_with_single_wildcard() {
        let source = r#"
Sub Test()
    Kill "C:\TEST?.TXT"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("TEST?.TXT"));
    }

    #[test]
    fn kill_with_variable() {
        let source = r"
Sub Test()
    Kill myFileName
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("myFileName"));
    }

    #[test]
    fn kill_with_app_path() {
        let source = r#"
Sub Test()
    Kill App.Path & "\temp.dat"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("App"));
    }

    #[test]
    fn kill_preserves_whitespace() {
        let source = "    Kill    \"C:\\\\file.txt\"    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Kill    \"C:\\\\file.txt\"    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
    }

    #[test]
    fn kill_with_comment() {
        let source = r#"
Sub Test()
    Kill "temp.txt" ' Delete temporary file
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn kill_in_if_statement() {
        let source = r"
Sub Test()
    If fileExists Then
        Kill fileName
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
    }

    #[test]
    fn kill_inline_if() {
        let source = r#"
Sub Test()
    If fileExists Then Kill "temp.dat"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
    }

    #[test]
    fn kill_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Kill "temp.txt"
    If Err.Number <> 0 Then
        MsgBox "Could not delete file"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
    }

    #[test]
    fn multiple_kill_statements() {
        let source = r#"
Sub Test()
    Kill "temp1.txt"
    Kill "temp2.txt"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let kill_count = debug.matches("KillStatement").count();
        assert_eq!(kill_count, 2);
    }

    #[test]
    fn kill_with_dir_function() {
        let source = r#"
Sub Test()
    fileName = Dir("C:\*.tmp")
    Do While fileName <> ""
        Kill "C:\" & fileName
        fileName = Dir
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("KillStatement"));
    }
}
