use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    // VB6 ChDrive statement syntax:
    // - ChDrive drive
    //
    // Changes the current drive.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
    pub(crate) fn parse_ch_drive_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ChDriveStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn chdrive_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDrive "C:"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_variable() {
        let source = r"
Sub Test()
    ChDrive myDrive
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_app_path() {
        let source = r"
Sub Test()
    ChDrive App.Path
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_left_function() {
        let source = r"
Sub Test()
    ChDrive Left(sInitDir, 1)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_in_if_statement() {
        let source = r"
Sub Test()
    If driveValid Then ChDrive newDrive
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_at_module_level() {
        let source = r#"
ChDrive "D:"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_comment() {
        let source = r"
Sub Test()
    ChDrive driveLetter ' Change to specified drive
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn chdrive_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDrive "C:"
    ChDrive "D:"
    ChDrive originalDrive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let chdrive_count = debug.matches("ChDriveStatement").count();
        assert_eq!(chdrive_count, 3, "Expected 3 ChDrive statements");
    }

    #[test]
    fn chdrive_in_multiline_if() {
        let source = r"
Sub Test()
    If driveExists Then
        ChDrive targetDrive
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_parentheses() {
        let source = r"
Sub Test()
    ChDrive (Left$(sInitDir, 1))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_expression() {
        let source = r"
Sub Test()
    ChDrive Left(theZtmPath, 1)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_and_chdir_together() {
        let source = r#"
Sub Test()
    ChDrive "C:"
    ChDir "C:\Windows"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }
}
