use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Parse a Put statement.
    ///
    /// VB6 Put statement syntax:
    /// - Put [#]filenumber, [recnumber], varname
    ///
    /// Writes data from a variable to a disk file.
    ///
    /// The Put statement syntax has these parts:
    ///
    /// | Part          | Description |
    /// |---------------|-------------|
    /// | filenumber    | Required. Any valid file number. |
    /// | recnumber     | Optional. Variant (Long). Record number (Random mode files) or byte number (Binary mode files) at which writing begins. |
    /// | varname       | Required. Valid variable name containing data to be written to disk. |
    ///
    /// Remarks:
    /// - Put is used with files opened in Binary or Random mode.
    /// - For files opened in Random mode, the record length specified in the Open statement determines the number of bytes written.
    /// - For files opened in Binary mode, Put writes any number of bytes.
    /// - The first record or byte in a file is at position 1, the second at position 2, and so on.
    /// - If you omit recnumber, the next record or byte following the last Put or Get statement (or pointed to by the last Seek function) is written.
    /// - You must include delimiting commas, for example: Put #1, , myVariable
    /// - For files opened in Random mode, the following rules apply:
    ///   * If the length of the data being written is less than the length specified in the Len clause, subsequent records on disk are aligned on record-length boundaries.
    ///   * The space between the end of one record and the beginning of the next is padded with the existing file contents.
    ///   * If the variable being written is a variable-length string, Put writes a 2-byte descriptor containing the string length and then writes the string data.
    /// - For files opened in Binary mode, all the Random rules apply, except:
    ///   * The Len clause in the Open statement has no effect.
    ///   * Put writes the data contiguously, with no padding between records.
    /// - Put statements usually mirror Get statements. That is, data written with Put is typically read with Get.
    ///
    /// Examples:
    /// ```vb
    /// Put #1, , myRecord
    /// Put #1, recordNumber, customerData
    /// Put fileNum, , buffer
    /// Put #1, filePosition, userData
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/put-statement)
    pub(crate) fn parse_put_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::PutStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Put statement tests
    #[test]
    fn put_simple() {
        let source = r"
Sub Test()
    Put #1, , myRecord
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("PutKeyword"));
    }

    #[test]
    fn put_at_module_level() {
        let source = "Put #1, , myData\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_record_number() {
        let source = r"
Sub Test()
    Put #1, recordNumber, customerData
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("recordNumber"));
    }

    #[test]
    fn put_with_file_variable() {
        let source = r"
Sub Test()
    Put fileNum, , buffer
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("fileNum"));
    }

    #[test]
    fn put_with_hash_symbol() {
        let source = r"
Sub Test()
    Put #fileNumber, position, data
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_preserves_whitespace() {
        let source = "    Put    #1  ,  ,  myVar    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Put    #1  ,  ,  myVar    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_comment() {
        let source = r"
Sub Test()
    Put #1, , myRecord ' Write next record
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn put_in_if_statement() {
        let source = r"
Sub Test()
    If dataReady Then
        Put #1, , myData
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_multiple_in_sequence() {
        let source = r"
Sub Test()
    Put #1, , record1
    Put #1, , record2
    Put #1, , record3
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let put_count = debug.matches("PutStatement").count();
        assert_eq!(put_count, 3);
    }

    #[test]
    fn put_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Put #1, , records(i)
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_udt() {
        let source = r"
Sub Test()
    Put #1, , employee
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("employee"));
    }

    #[test]
    fn put_binary_data() {
        let source = r"
Sub Test()
    Put #1, bytePosition, buffer()
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("bytePosition"));
    }

    #[test]
    fn put_with_seek_position() {
        let source = r"
Sub Test()
    Put #1, Seek(1), myData
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_inline_if() {
        let source = r"
Sub Test()
    If writeFlag Then Put #1, , record
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Put #1, , myRecord
    If Err.Number <> 0 Then
        MsgBox "Error writing record"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_after_get() {
        let source = r"
Sub Test()
    Get #1, recordNum, myRecord
    ' Modify the record
    Put #1, recordNum, myRecord
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn put_with_explicit_position() {
        let source = r"
Sub Test()
    Put #1, 100, headerData
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_calculated_position() {
        let source = r"
Sub Test()
    Put #1, (recordNum - 1) * recordLength + 1, myData
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_array_element() {
        let source = r"
Sub Test()
    Put #1, , dataArray(index)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("dataArray"));
    }

    #[test]
    fn put_object_property() {
        let source = r"
Sub Test()
    Put #1, , myObject.Data
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_multiline_if() {
        let source = r"
Sub Test()
    If needsWrite Then
        Put #1, recordPos, recordData
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_in_select_case() {
        let source = r"
Sub Test()
    Select Case recordType
        Case 1
            Put #1, , type1Record
        Case 2
            Put #1, , type2Record
    End Select
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let put_count = debug.matches("PutStatement").count();
        assert_eq!(put_count, 2);
    }

    #[test]
    fn put_string_variable() {
        let source = r"
Sub Test()
    Put #1, , userName
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("userName"));
    }

    #[test]
    fn put_numeric_literal_position() {
        let source = r"
Sub Test()
    Put #1, 1, headerRecord
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_do_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Get #1, , inRecord
        Put #2, , outRecord
    Loop
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn put_random_access_file() {
        let source = r#"
Sub Test()
    Open "data.dat" For Random As #1 Len = Len(myRecord)
    Put #1, recordNumber, myRecord
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
        assert!(debug.contains("OpenStatement"));
    }

    #[test]
    fn put_binary_file() {
        let source = r#"
Sub Test()
    Open "binary.bin" For Binary As #1
    Put #1, , byteArray()
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }
}
