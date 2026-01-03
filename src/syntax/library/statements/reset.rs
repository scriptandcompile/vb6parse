use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Reset statement syntax:
    // - Reset
    //
    // Closes all disk files opened using the Open statement.
    //
    // The Reset statement closes all active files opened by the Open statement
    // and writes the contents of all file buffers to disk.
    //
    // Use Reset to ensure all file data is written to disk before ending your program.
    // This is particularly important in programs that may terminate abnormally.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/reset-statement)
    pub(crate) fn parse_reset_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ResetStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    // Reset statement tests
    #[test]
    fn reset_simple() {
        let source = r"
Sub Test()
    Reset
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
        assert!(debug.contains("ResetKeyword"));
    }

    #[test]
    fn reset_at_module_level() {
        let source = "Reset\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_in_if_statement() {
        let source = r"
Sub CleanupFiles()
    If CloseAllFiles Then
        Reset
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_preserves_whitespace() {
        let source = "    Reset    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Reset    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_with_comment() {
        let source = r"
Sub Test()
    Reset ' Close all open files
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn reset_inline_if() {
        let source = r"
Sub Test()
    If shouldClose Then Reset
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Reset
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Reset
    If Err.Number <> 0 Then
        MsgBox "Error closing files"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn multiple_reset_statements() {
        let source = r"
Sub Test()
    Reset
    DoSomething
    Reset
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let reset_count = debug.matches("ResetStatement").count();
        assert_eq!(reset_count, 2);
    }

    #[test]
    fn reset_after_file_operations() {
        let source = r#"
Sub Test()
    Open "test.txt" For Output As #1
    Print #1, "data"
    Close #1
    Reset
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }
}
