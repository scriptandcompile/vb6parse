use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Open statement syntax:
    // - Open pathname For mode [Access access] [lock] As [#]filenumber [Len=reclength]
    //
    // Enables input/output (I/O) to a file.
    //
    // The Open statement syntax has these parts:
    //
    // | Part       | Description |
    // |------------|-------------|
    // | pathname   | Required. String expression that specifies a file name — may include directory or folder, and drive. |
    // | mode       | Required. Keyword specifying the file mode: Append, Binary, Input, Output, or Random. If unspecified, the file is opened for Random access. |
    // | access     | Optional. Keyword specifying the operations permitted on the open file: Read, Write, or Read Write. |
    // | lock       | Optional. Keyword specifying the operations restricted on the open file by other processes: Shared, Lock Read, Lock Write, and Lock Read Write. |
    // | filenumber | Required. A valid file number in the range 1 to 511, inclusive. Use the FreeFile function to obtain the next available file number. |
    // | reclength  | Optional. Number less than or equal to 32,767 (bytes). For files opened for random access, this value is the record length. For sequential files, this value is the number of characters buffered. |
    //
    // Remarks:
    // - You must open a file before any I/O operation can be performed on it.
    // - If pathname specifies a file that doesn't exist, it is created when a file is opened for Append, Binary, Output, or Random modes.
    // - If the file is already opened by another process and the specified type of access is not allowed, the Open operation fails and an error occurs.
    // - The Len clause is ignored if mode is Binary.
    // - In Binary, Input, and Random modes, you can open a file using a different file number without first closing the file. In Append and Output modes, you must close a file before opening it with a different file number.
    //
    // Examples:
    // ```vb
    // ' Open for input
    // Open "TESTFILE" For Input As #1
    //
    // ' Open for output
    // Open "TESTFILE" For Output As #1
    //
    // ' Open for append
    // Open "TESTFILE" For Append As #1
    //
    // ' Open for binary
    // Open "TESTFILE" For Binary As #1
    //
    // ' Open for random with record length
    // Open "TESTFILE" For Random As #1 Len = 512
    //
    // ' Open with access control
    // Open "TESTFILE" For Input Access Read As #1
    //
    // ' Open with locking
    // Open "TESTFILE" For Binary Lock Read Write As #1
    //
    // ' Open with variable
    // Dim fileNum As Integer
    // fileNum = FreeFile
    // Open fileName For Input As fileNum
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    pub(crate) fn parse_open_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::OpenStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn open_for_input() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Input As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_for_output() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Output As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_for_append() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Append As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_for_binary() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Binary As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_for_random() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Random As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_len() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Random As #1 Len = 512
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_access_read() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Input Access Read As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_access_write() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Output Access Write As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_access_read_write() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Binary Access Read Write As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_lock_read() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Binary Lock Read As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_lock_write() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Binary Lock Write As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_lock_read_write() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Binary Lock Read Write As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_shared() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Input Shared As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_variable_filename() {
        let source = r#"
Sub Test()
    Dim fileName As String
    fileName = "test.txt"
    Open fileName For Input As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_freefile() {
        let source = r#"
Sub Test()
    Dim fileNum As Integer
    fileNum = FreeFile
    Open "TESTFILE" For Input As fileNum
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_without_hash() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Input As 1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_path() {
        let source = r#"
Sub Test()
    Open "C:\Temp\TESTFILE.txt" For Output As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_preserves_whitespace() {
        let source = r#"
Sub Test()
    Open   "TESTFILE"   For   Input   As   #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_in_if_statement() {
        let source = r#"
Sub Test()
    If fileExists Then
        Open "TESTFILE" For Input As #1
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_inline_if() {
        let source = r#"
Sub Test()
    If needsFile Then Open "TESTFILE" For Input As #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_open_statements() {
        let source = r#"
Sub Test()
    Open "FILE1.txt" For Input As #1
    Open "FILE2.txt" For Output As #2
    Open "FILE3.txt" For Append As #3
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Open "TESTFILE" For Input As #1
    If Err.Number <> 0 Then MsgBox "Error opening file"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_at_module_level() {
        let source = r#"Open "TESTFILE" For Input As #1"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_with_comment() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Input As #1 ' Open file for reading
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn open_complete_syntax() {
        let source = r#"
Sub Test()
    Open "TESTFILE" For Random Access Read Write Lock Read Write As #1 Len = 512
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/open");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
