//! # Name Statement
//!
//! Renames a disk file, directory, or folder.
//!
//! ## Syntax
//!
//! ```vb
//! Name oldpathname As newpathname
//! ```
//!
//! - `oldpathname`: Required. String expression that specifies the existing file name and location. May include directory or folder, and drive.
//! - `newpathname`: Required. String expression that specifies the new file name and location. May include directory or folder, and drive.
//!   Cannot specify a different drive from the one specified in `oldpathname`.
//!
//! ## Remarks
//!
//! - The `Name` statement renames a file and moves it to a different directory or folder, if necessary
//! - `Name` can move a file across directories or folders, but both `oldpathname` and `newpathname` must be on the same drive
//! - Using `Name` on an open file produces an error. You must close an open file before renaming it
//! - `Name` arguments can include relative or absolute paths
//! - The `Name` statement can also rename directories or folders
//! - If `newpathname` already exists, an error occurs
//! - Wildcard characters (* and ?) are not allowed in either `oldpathname` or `newpathname`
//!
//! ## Examples
//!
//! ```vb
//! ' Rename a file
//! Name "OLDFILE.TXT" As "NEWFILE.TXT"
//!
//! ' Move and rename a file
//! Name "C:\Data\Report.doc" As "C:\Archive\OldReport.doc"
//!
//! ' Rename a directory
//! Name "C:\OldFolder" As "C:\NewFolder"
//!
//! ' Move file to different directory (same drive)
//! Name "C:\Temp\Test.dat" As "C:\Data\Test.dat"
//!
//! ' Using variables
//! Dim oldName As String, newName As String
//! oldName = "File1.txt"
//! newName = "File2.txt"
//! Name oldName As newName
//! ```
//!
//! ## Reference
//!
//! [Name Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/name-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Name statement.
    pub(crate) fn parse_name_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::NameStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn name_simple() {
        let source = r#"
Sub Test()
    Name "OLDFILE.TXT" As "NEWFILE.TXT"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_at_module_level() {
        let source = r#"Name "old.txt" As "new.txt""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_with_full_paths() {
        let source = r#"
Sub Test()
    Name "C:\Data\Report.doc" As "C:\Archive\OldReport.doc"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_with_variables() {
        let source = r"
Sub Test()
    Name oldName As newName
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_preserves_whitespace() {
        let source = "    Name    \"old.txt\"    As    \"new.txt\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_with_comment() {
        let source = r#"
Sub Test()
    Name "temp.dat" As "backup.dat" ' Rename temp file
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_in_if_statement() {
        let source = r#"
Sub Test()
    If fileExists Then
        Name "old.log" As "archive.log"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_inline_if() {
        let source = r"
Sub Test()
    If needsRename Then Name oldFile As newFile
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_name_statements() {
        let source = r#"
Sub RenameFiles()
    Name "File1.txt" As "Backup1.txt"
    Name "File2.txt" As "Backup2.txt"
    Name "File3.txt" As "Backup3.txt"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_rename_directory() {
        let source = r#"
Sub Test()
    Name "C:\OldFolder" As "C:\NewFolder"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_with_concatenation() {
        let source = r#"
Sub Test()
    Name basePath & "old.dat" As basePath & "new.dat"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn name_move_file() {
        let source = r#"
Sub Test()
    Name "C:\Temp\Test.dat" As "C:\Data\Test.dat"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/library/statements/name");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
