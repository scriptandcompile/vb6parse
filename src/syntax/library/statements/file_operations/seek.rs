//! # Seek Statement
//!
//! Sets the position for the next read or write operation in a file opened using the Open statement.
//!
//! ## Syntax
//!
//! ```vb
//! Seek [#]filenumber, position
//! ```
//!
//! ## Parts
//!
//! - **filenumber**: Required. Any valid file number. The number sign (#) is optional but commonly included for clarity.
//! - **position**: Required. Number in the range 1 to 2,147,483,647 (equivalent to 2^31 - 1), indicating where the next read or write should occur.
//!
//! ## Remarks
//!
//! - **File Position**: The Seek statement sets the byte position in a file where the next Input, Output, Get, or Put operation will occur.
//! - **Position Numbering**: File positions are numbered beginning with 1 (the first byte in the file is at position 1, not 0).
//! - **Random Access Files**: For Random mode files, the position parameter specifies a record number rather than a byte position.
//! - **Sequential Files**: For files opened in Input, Output, or Append mode, position specifies the byte position.
//! - **Binary Files**: For Binary mode files, position specifies the byte position.
//! - **Seek Function**: Use the Seek function (without arguments except file number) to return the current file position.
//! - **EOF Behavior**: Setting the position beyond the end of the file doesn't immediately extend the file, but writing to that position will.
//! - **Position Range**: The position must be a positive Long value (1 to 2,147,483,647).
//! - **File Number**: The file must be opened before using Seek.
//!
//! ## Position Interpretation by File Mode
//!
//! | File Mode | Position Represents |
//! |-----------|-------------------|
//! | Random    | Record number (1-based) |
//! | Binary    | Byte position (1-based) |
//! | Input     | Byte position (1-based) |
//! | Output    | Byte position (1-based) |
//! | Append    | Byte position (1-based) |
//!
//! ## Examples
//!
//! ### Seek to Beginning of File
//!
//! ```vb
//! Open "DATA.TXT" For Binary As #1
//! Seek #1, 1   ' Position at first byte
//! ' Read or write operations
//! Close #1
//! ```
//!
//! ### Seek to Specific Byte Position
//!
//! ```vb
//! Open "BINARY.DAT" For Binary As #1
//! Seek #1, 100   ' Position at byte 100
//! Get #1, , myData
//! Close #1
//! ```
//!
//! ### Seek to Specific Record in Random File
//!
//! ```vb
//! Type Employee
//!     ID As Integer
//!     Name As String * 30
//! End Type
//!
//! Dim emp As Employee
//! Open "EMPLOYEE.DAT" For Random As #1 Len = Len(emp)
//! Seek #1, 5   ' Position at record 5
//! Get #1, , emp
//! Close #1
//! ```
//!
//! ### Seek Based on Calculation
//!
//! ```vb
//! Dim recordNumber As Long
//! recordNumber = 10
//! Seek #1, recordNumber
//! ```
//!
//! ### Using Seek with Loop
//!
//! ```vb
//! Open "DATA.BIN" For Binary As #1
//! For i = 1 To 100 Step 10
//!     Seek #1, i
//!     Put #1, , dataArray(i)
//! Next i
//! Close #1
//! ```
//!
//! ### Seek to End of File
//!
//! ```vb
//! Open "APPEND.TXT" For Binary As #1
//! Seek #1, LOF(1) + 1   ' Position after last byte
//! Put #1, , newData
//! Close #1
//! ```
//!
//! ### Combined with Seek Function
//!
//! ```vb
//! Open "DATA.TXT" For Binary As #1
//! currentPos = Seek(1)      ' Get current position
//! Seek #1, currentPos + 50  ' Move 50 bytes forward
//! Close #1
//! ```
//!
//! ### Rewind File
//!
//! ```vb
//! Sub RewindFile(fileNum As Integer)
//!     Seek fileNum, 1  ' Return to beginning
//! End Sub
//! ```
//!
//! ### Seek in Random Access Processing
//!
//! ```vb
//! Type Product
//!     Code As String * 10
//!     Price As Double
//! End Type
//!
//! Dim prod As Product
//! Dim recordNum As Long
//!
//! Open "PRODUCTS.DAT" For Random As #1 Len = Len(prod)
//! recordNum = 25
//! Seek #1, recordNum
//! Get #1, , prod
//! prod.Price = prod.Price * 1.1  ' Increase price by 10%
//! Seek #1, recordNum
//! Put #1, , prod
//! Close #1
//! ```
//!
//! ## Common Errors
//!
//! - **Error 52**: Bad file name or number - file not open or invalid file number
//! - **Error 63**: Bad record number - position is less than 1 or exceeds valid range
//! - **Error 5**: Invalid procedure call - negative or zero position value
//!
//! ## Performance Tips
//!
//! - For sequential reading/writing, you generally don't need Seek as the file pointer advances automatically.
//! - Use Seek when you need random access to specific parts of a file.
//! - Combining Seek with the Seek function allows you to save and restore file positions.
//! - For large files, seeking to specific positions is much faster than reading sequentially.
//!
//! ## See Also
//!
//! - `Seek` function (returns current file position)
//! - `Get` statement (read data from file)
//! - `Put` statement (write data to file)
//! - `Open` statement (open files)
//! - `LOF` function (length of file)
//! - `Loc` function (current position in file)
//!
//! ## References
//!
//! - [Seek Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/seek-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Seek statement.
    pub(crate) fn parse_seek_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::SeekStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn seek_simple() {
        let source = r"
Sub Test()
    Seek #1, 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_at_module_level() {
        let source = "Seek #1, 100\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_without_hash() {
        let source = r"
Sub Test()
    Seek 1, 50
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_variable() {
        let source = r"
Sub Test()
    Seek #1, position
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_to_beginning() {
        let source = r"
Sub Test()
    Seek #fileNum, 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_expression() {
        let source = r"
Sub Test()
    Seek #1, recordNum * recordSize
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_lof() {
        let source = r"
Sub Test()
    Seek #1, LOF(1) + 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_seek_function() {
        let source = r"
Sub Test()
    Seek #1, Seek(1) + 100
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_inside_if_statement() {
        let source = r"
If needSeek Then
    Seek #1, targetPosition
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_inside_loop() {
        let source = r"
For i = 1 To 100
    Seek #1, i * 10
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_comment() {
        let source = r"
Sub Test()
    Seek #1, 1  ' Rewind to start
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_preserves_whitespace() {
        let source = "Seek   #1  ,   100\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_multiple_files() {
        let source = r"
Sub Test()
    Seek #1, pos1
    Seek #2, pos2
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_addition() {
        let source = r"
Sub Test()
    Seek #1, currentPos + offset
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_select_case() {
        let source = r"
Select Case action
    Case 1
        Seek #1, 100
    Case 2
        Seek #1, 200
End Select
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_random_access() {
        let source = r"
Sub Test()
    Seek #1, recordNumber
    Get #1, , myRecord
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_binary_mode() {
        let source = r"
Sub Test()
    Seek #1, bytePosition
    Put #1, , dataBytes
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_sub() {
        let source = r"
Sub RewindFile(fileNum As Integer)
    Seek fileNum, 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_function() {
        let source = r#"
Function GetRecordAt(fileNum As Integer, pos As Long) As String
    Seek fileNum, pos
    GetRecordAt = ""
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_parentheses() {
        let source = r"
Sub Test()
    Seek #1, (recordNum - 1) * recordLen + 1
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_error_handler() {
        let source = r#"
On Error Resume Next
Seek #1, targetPos
If Err.Number <> 0 Then
    MsgBox "Seek failed"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_multiple_on_same_line() {
        let source = "Seek #1, 100: Seek #2, 200\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_with_block() {
        let source = r"
With fileData
    Seek .fileNum, .position
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_large_number() {
        let source = r"
Sub Test()
    Seek #1, 2147483647
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_in_class_module() {
        let source = r"
Private fileNumber As Integer

Public Sub SetPosition(pos As Long)
    Seek fileNumber, pos
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_line_continuation() {
        let source = r"
Sub Test()
    Seek #1, _
        targetPosition
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_double_seek_pattern() {
        let source = r"
Sub Test()
    Seek #1, recordNum
    Get #1, , data
    Seek #1, recordNum
    Put #1, , data
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_with_clng() {
        let source = r"
Sub Test()
    Seek #1, CLng(position)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn seek_forward_backward() {
        let source = r"
Sub Test()
    currentPos = Seek(1)
    Seek #1, currentPos + 10
    Seek #1, currentPos - 5
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/seek");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
