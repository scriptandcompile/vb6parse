use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    // VB6 Get statement syntax:
    // - Get [#]filenumber, [recnumber], varname
    //
    // Reads data from an open disk file into a variable.
    //
    // The Get statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recnumber     | Optional. Variant (Long). Record number (Random mode files) or byte number (Binary mode files) at which reading begins. |
    // | varname       | Required. Valid variable name into which data is read. |
    //
    // Remarks:
    // - Get is used with files opened in Binary or Random mode.
    // - For files opened in Random mode, the record length specified in the Open statement determines the number of bytes read.
    // - For files opened in Binary mode, Get reads any number of bytes.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you omit recnumber, the next record or byte following the last Get or Put statement (or pointed to by the last Seek function) is read.
    // - You must include delimiting commas, for example: Get #1, , myVariable
    // - For files opened in Random mode, the following rules apply:
    //   * If the length of the data being read is less than the length specified in the Len clause, subsequent records on disk are aligned on record-length boundaries.
    //   * The space between the end of one record and the beginning of the next is padded with existing file contents.
    //   * If the variable being read is a variable-length string, Get reads a 2-byte descriptor containing the string length and then reads the string data.
    // - For files opened in Binary mode, all the Random rules apply, except:
    //   * The Len clause in the Open statement has no effect.
    //   * Get reads the data contiguously, with no padding between records.
    //
    // Examples:
    // ```vb
    // Get #1, , myRecord
    // Get #1, recordNumber, customerData
    // Get fileNum, , buffer
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/get-statement)
    pub(crate) fn parse_get_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::GetStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Get statement tests
    #[test]
    fn get_simple() {
        let source = r#"
Sub Test()
    Get #1, , myRecord
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("GetKeyword"));
    }

    #[test]
    fn get_at_module_level() {
        let source = "Get #1, , myData\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_record_number() {
        let source = r#"
Sub Test()
    Get #1, recordNumber, customerData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("recordNumber"));
    }

    #[test]
    fn get_with_file_variable() {
        let source = r#"
Sub Test()
    Get fileNum, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("fileNum"));
    }

    #[test]
    fn get_with_hash_symbol() {
        let source = r#"
Sub Test()
    Get #fileNumber, position, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_preserves_whitespace() {
        let source = "    Get    #1  ,  ,  myVar    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Get    #1  ,  ,  myVar    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_comment() {
        let source = r#"
Sub Test()
    Get #1, , myRecord ' Read next record
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn get_in_if_statement() {
        let source = r#"
Sub Test()
    If Not EOF(1) Then
        Get #1, , myData
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_inline_if() {
        let source = r#"
Sub Test()
    If hasData Then Get #1, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Get #1, , myRecord
    If Err.Number <> 0 Then
        MsgBox "Error reading file"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_in_loop() {
        let source = r#"
Sub Test()
    Do While Not EOF(1)
        Get #1, , myRecord
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn multiple_get_statements() {
        let source = r#"
Sub Test()
    Get #1, , record1
    Get #1, , record2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let get_count = debug.matches("GetStatement").count();
        assert_eq!(get_count, 2);
    }

    #[test]
    fn get_binary_mode() {
        let source = r#"
Sub Test()
    Dim buffer As String * 512
    Get #1, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }
}
