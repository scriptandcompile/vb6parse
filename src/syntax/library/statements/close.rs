use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Close statement syntax:
    // - Close [filenumberlist]
    //
    // Closes input or output files opened using the Open statement.
    //
    // filenumberlist: Optional. One or more file numbers using the syntax:
    // [[#]filenumber] [, [#]filenumber] ...
    //
    // If filenumberlist is omitted, all active files opened by the Open statement are closed.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/close-statement)
    pub(crate) fn parse_close_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::CloseStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn close_all_files() {
        let source = r"
Sub Test()
    Close
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_single_file() {
        let source = r"
Sub Test()
    Close #1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_single_file_without_hash() {
        let source = r"
Sub Test()
    Close 1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_multiple_files() {
        let source = r"
Sub Test()
    Close #1, #2, #3
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_with_variable() {
        let source = r"
Sub Test()
    Close fileNum
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_with_hash_variable() {
        let source = r"
Sub Test()
    Close #fileNum
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_multiple_files_mixed() {
        let source = r"
Sub Test()
    Close #1, fileNum2, #3
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_preserves_whitespace() {
        let source = r"
Sub Test()
    Close   #1  ,  #2
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_close_statements() {
        let source = r"
Sub Test()
    Close #1
    Close #2
    Close
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("CloseStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn close_in_if_statement() {
        let source = r"
Sub Test()
    If fileOpen Then
        Close #1
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn close_inline_if() {
        let source = r"
Sub Test()
    If fileOpen Then Close #fileNum
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }

    #[test]
    fn close_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Close #1
    If Err Then MsgBox "Error closing file"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }

    #[test]
    fn close_at_module_level() {
        let source = r"
Close #1
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }
}
