use crate::language::Token;
use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Line Input statement syntax:
    // - Line Input #filenumber, varname
    //
    // Reads a single line from an open sequential file and assigns it to a String variable.
    //
    // The Line Input # statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | varname       | Required. Valid String or Variant variable name. |
    //
    // Remarks:
    // - Data read with Line Input # is usually written to a file with Print #.
    // - The Line Input # statement reads from a file one character at a time until it encounters
    //   a carriage return (Chr(13)) or carriage return–linefeed (Chr(13) + Chr(10)) sequence.
    // - Carriage return–linefeed sequences are skipped rather than appended to the character string.
    // - Line Input # is useful for reading text files that have been created in a text editor or
    //   with the Print # statement.
    // - Unlike Input #, Line Input # doesn't parse the data as it's read – you get the entire line as-is.
    // - If end of file is reached before reading a complete line, an error occurs.
    //
    // Examples:
    // ```vb
    // Line Input #1, textLine
    // Line Input #fileNum, dataBuffer
    // Line Input #1, myString
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/line-input-statement)
    pub(crate) fn parse_line_input_statement(&mut self) {
        // if we are now parsing a Line Input statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::LineInputStatement.to_raw());

        // Consume "Line" keyword
        self.consume_token();

        // Consume "Input" keyword (should be next)
        if self.at_token(Token::InputKeyword) {
            self.consume_token();
        }

        // Consume everything until newline
        // This includes: "#", filenumber, ",", varname
        while !self.is_at_end() && !self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LineInputStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn line_input_simple() {
        let source = r"
Sub Test()
    Line Input #1, textLine
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_at_module_level() {
        let source = "Line Input #1, myLine\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_with_file_variable() {
        let source = r"
Sub Test()
    Line Input #fileNum, buffer
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_preserves_whitespace() {
        let source = "    Line    Input    #1  ,  myString    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_with_comment() {
        let source = r"
Sub Test()
    Line Input #1, textLine ' Read one line
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_in_if_statement() {
        let source = r"
Sub Test()
    If Not EOF(1) Then
        Line Input #1, currentLine
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_inline_if() {
        let source = r"
Sub Test()
    If hasData Then Line Input #1, nextLine
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_in_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Line Input #1, textLine
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_line_input_statements() {
        let source = r"
Sub Test()
    Line Input #1, line1
    Line Input #1, line2
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_reading_text_file() {
        let source = r#"
Sub Test()
    Open "data.txt" For Input As #1
    Line Input #1, headerLine
    Close #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Line Input #1, textLine
    If Err.Number <> 0 Then
        MsgBox "Error reading file"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_vs_input() {
        let source = r"
Sub Test()
    Line Input #1, wholeLine
    Input #1, parsedData
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn line_input_string_variable() {
        let source = r"
Sub Test()
    Dim myText As String
    Line Input #1, myText
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../../snapshots/syntax/library/statements/line_input");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
