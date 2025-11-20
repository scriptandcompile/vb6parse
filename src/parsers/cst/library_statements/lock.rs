use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    // VB6 Lock statement syntax:
    // - Lock [#]filenumber[, recordrange]
    //
    // Controls access to all or part of an open file.
    //
    // The Lock statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recordrange   | Optional. Range of records to lock. Can be: record, start To end, or omitted for entire file. |
    //
    // Remarks:
    // - Lock and Unlock are used in environments where multiple processes might need access to the same file.
    // - Lock and Unlock statements are always used in pairs.
    // - The Lock statement locks all or part of a file opened using the Open statement.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you specify just one record number, only that record is locked.
    // - If you specify a range, all records in that range are locked.
    // - For files opened in Binary, Input, or Output mode, Lock always locks the entire file,
    //   regardless of the recordrange argument.
    // - For files opened in Random mode, Lock locks the specified record or range of records.
    // - Locked portions of a file can't be accessed by other processes until unlocked with Unlock.
    // - Use Unlock to remove the lock from a portion of a file.
    //
    // Examples:
    // ```vb
    // Lock #1
    // Lock #1, 5
    // Lock #1, 10 To 20
    // Lock fileNum, recordNum
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/lock-statement)
    pub(crate) fn parse_lock_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::LockStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Lock statement tests
    #[test]
    fn lock_simple() {
        let source = r#"
Sub Test()
    Lock #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
        assert!(debug.contains("LockKeyword"));
    }

    #[test]
    fn lock_at_module_level() {
        let source = "Lock #1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_entire_file() {
        let source = r#"
Sub Test()
    Lock #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_single_record() {
        let source = r#"
Sub Test()
    Lock #1, 5
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_record_range() {
        let source = r#"
Sub Test()
    Lock #1, 10 To 20
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
        assert!(debug.contains("ToKeyword"));
    }

    #[test]
    fn lock_with_variable() {
        let source = r#"
Sub Test()
    Lock fileNum, recordNum
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
        assert!(debug.contains("fileNum"));
        assert!(debug.contains("recordNum"));
    }

    #[test]
    fn lock_preserves_whitespace() {
        let source = "    Lock    #1  ,  5    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Lock    #1  ,  5    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_with_comment() {
        let source = r#"
Sub Test()
    Lock #1, 5 ' Lock record 5
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn lock_in_if_statement() {
        let source = r#"
Sub Test()
    If needsLock Then
        Lock #1, currentRecord
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_inline_if() {
        let source = r#"
Sub Test()
    If multiUser Then Lock #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_unlock_pair() {
        let source = r#"
Sub Test()
    Lock #1, 5
    ' Do work
    Unlock #1, 5
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Lock #1, recordNum
    If Err.Number <> 0 Then
        MsgBox "Could not lock record"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_for_shared_file_access() {
        let source = r#"
Sub Test()
    Open "data.dat" For Random As #1 Shared
    Lock #1, 10
    ' Write to record 10
    Put #1, 10, myData
    Unlock #1, 10
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }
}
