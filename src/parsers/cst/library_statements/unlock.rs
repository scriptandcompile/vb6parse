use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    // VB6 Unlock statement syntax:
    // - Unlock [#]filenumber[, recordrange]
    //
    // Removes access restrictions on all or part of an open file.
    //
    // The Unlock statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recordrange   | Optional. Range of records to unlock. Can be: record, start To end, or omitted for entire file. |
    //
    // Remarks:
    // - Unlock is used to remove locks placed on a file with the Lock statement.
    // - The Unlock statement allows other processes to access the unlocked portions of the file.
    // - The arguments to Unlock must exactly match those used with the corresponding Lock statement.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you specify just one record number, only that record is unlocked.
    // - If you specify a range, all records in that range are unlocked.
    // - For files opened in Binary, Input, or Output mode, Unlock always unlocks the entire file,
    //   regardless of the recordrange argument.
    // - For files opened in Random mode, Unlock unlocks the specified record or range of records.
    // - Each Lock statement must have a corresponding Unlock statement with the same file number
    //   and record range.
    //
    // Examples:
    // ```vb
    // Unlock #1
    // Unlock #1, 5
    // Unlock #1, 10 To 20
    // Unlock fileNum, recordNum
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unlock-statement)
    pub(crate) fn parse_unlock_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::UnlockStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Unlock statement tests
    #[test]
    fn unlock_simple() {
        let source = r"
Sub Test()
    Unlock #1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
        assert!(debug.contains("UnlockKeyword"));
    }

    #[test]
    fn unlock_at_module_level() {
        let source = "Unlock #1\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_entire_file() {
        let source = r"
Sub Test()
    Unlock #1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_single_record() {
        let source = r"
Sub Test()
    Unlock #1, 5
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_record_range() {
        let source = r"
Sub Test()
    Unlock #1, 10 To 20
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
        assert!(debug.contains("ToKeyword"));
    }

    #[test]
    fn unlock_with_variable() {
        let source = r"
Sub Test()
    Unlock fileNum, recordNum
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
        assert!(debug.contains("fileNum"));
        assert!(debug.contains("recordNum"));
    }

    #[test]
    fn unlock_preserves_whitespace() {
        let source = "    Unlock    #1  ,  5    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Unlock    #1  ,  5    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_with_comment() {
        let source = r"
Sub Test()
    Unlock #1, 5 ' Release record lock
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn unlock_in_if_statement() {
        let source = r"
Sub Test()
    If isDone Then
        Unlock #1, currentRecord
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_inline_if() {
        let source = r"
Sub Test()
    If finished Then Unlock #1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn lock_unlock_matching_pair() {
        let source = r"
Sub Test()
    Lock #1, 5
    myData.Value = 100
    Put #1, 5, myData
    Unlock #1, 5
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Unlock #1, recordNum
    If Err.Number <> 0 Then
        MsgBox "Could not unlock record"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("UnlockStatement"));
    }

    #[test]
    fn unlock_in_finally_block() {
        let source = r#"
Sub Test()
    Lock #1, recordNum
    On Error GoTo ErrorHandler
    ' Do work
    Put #1, recordNum, myData
    Unlock #1, recordNum
    Exit Sub
ErrorHandler:
    Unlock #1, recordNum
    MsgBox "Error occurred"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("UnlockStatement").count();
        assert_eq!(count, 2);
    }
}
