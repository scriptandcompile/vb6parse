use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 ChDir statement syntax:
    // - ChDir path
    //
    // Changes the current directory or folder.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
    pub(crate) fn parse_ch_dir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ChDirStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn chdir_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_variable() {
        let source = r"
Sub Test()
    ChDir myPath
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_app_path() {
        let source = r"
Sub Test()
    ChDir App.Path
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_expression() {
        let source = r#"
Sub Test()
    ChDir GetPath() & "\subdir"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_in_if_statement() {
        let source = r"
Sub Test()
    If dirExists Then ChDir newPath
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_at_module_level() {
        let source = r#"
ChDir "C:\Temp"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_comment() {
        let source = r"
Sub Test()
    ChDir basePath ' Change to base directory
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn chdir_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
    ChDir "C:\Temp"
    ChDir originalPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let chdir_count = debug.matches("ChDirStatement").count();
        assert_eq!(chdir_count, 3, "Expected 3 ChDir statements");
    }

    #[test]
    fn chdir_in_multiline_if() {
        let source = r"
Sub Test()
    If pathValid Then
        ChDir newPath
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_parentheses() {
        let source = r"
Sub Test()
    ChDir (basePath)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_parentheses_without_space() {
        let source = r"
Sub Test()
    ChDir(basePath)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }
}
