use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    // VB6 Print # statement syntax:
    // - Print #filenumber, [outputlist]
    //
    // Writes display-formatted data to a sequential file.
    //
    // The Print # statement syntax has these parts:
    //
    // | Part        | Description |
    // |-------------|-------------|
    // | filenumber  | Required. Any valid file number. |
    // | outputlist  | Optional. Expression or list of expressions to print. |
    //
    // Remarks:
    // - Data written with Print # is usually read from a file with Line Input # or Input.
    // - If you omit outputlist and include only a list separator after filenumber, a blank line is printed to the file.
    // - Multiple expressions can be separated with either a space or a semicolon.
    // - A space has the same effect as a semicolon.
    // - For Boolean data, either True or False is printed.
    // - The True and False keywords are not translated, regardless of locale.
    // - Date data is written to the file using the standard short date format recognized by your system.
    // - When either the date or the time component is missing or zero, only the part provided gets written to the file.
    // - Nothing is written to the file if outputlist data is Empty. However, if outputlist data is Null, Null is output to the file.
    // - For error data, the output appears as Error errorcode. The Error keyword is not translated, regardless of locale.
    // - All data written to the file using Print # is internationally aware; that is, the data is properly formatted using the appropriate decimal separator and thousands separator.
    // - When data is written to a file, several universal assumptions are followed:
    //   * Numeric data is always written using the period as the decimal separator.
    //   * For numeric data, a leading space is always reserved for the sign of the number.
    //   * A trailing space is included after each number.
    // - Unlike the Print method, the Print # statement doesn't insert commas or spaces between items as they are written to the file.
    // - When you use the Print # statement, you insert explicit delimiters in your output list when you want to add commas or spaces.
    // - The Print # statement usually writes Variant data to a file the same way it writes other data types.
    // - However, there are some exceptions:
    //   * If the data being written is a Variant of VarType vbError, an error message string is not written to the file.
    //   * Only the word Error and the error code are written.
    //   * If the data being written is a Variant of VarType vbEmpty, nothing is written to the file.
    //
    // Examples:
    // ```vb
    // ' Basic usage
    // Print #1, "Hello World"
    //
    // ' Multiple items
    // Print #1, x, y, z
    //
    // ' With semicolon separator
    // Print #1, "Name: "; userName; " Age: "; userAge
    //
    // ' Blank line
    // Print #1,
    //
    // ' Variable file number
    // Dim fileNum As Integer
    // fileNum = FreeFile
    // Print #fileNum, data
    //
    // ' Complex expressions
    // Print #1, Format$(Now, "yyyy-mm-dd"), totalAmount
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/print-statement)
    pub(crate) fn parse_print_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::PrintStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn print_basic() {
        let source = r#"
Sub Test()
    Print #1, "Hello World"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
    }

    #[test]
    fn print_multiple_items() {
        let source = r#"
Sub Test()
    Print #1, x, y, z
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("x"));
        assert!(debug.contains("y"));
        assert!(debug.contains("z"));
    }

    #[test]
    fn print_with_semicolon() {
        let source = r#"
Sub Test()
    Print #1, "Name: "; userName; " Age: "; userAge
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("userName"));
        assert!(debug.contains("userAge"));
    }

    #[test]
    fn print_blank_line() {
        let source = r#"
Sub Test()
    Print #1,
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
    }

    #[test]
    fn print_variable_file_number() {
        let source = r#"
Sub Test()
    Dim fileNum As Integer
    fileNum = FreeFile
    Print #fileNum, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("fileNum"));
        assert!(debug.contains("data"));
    }

    #[test]
    fn print_complex_expressions() {
        let source = r#"
Sub Test()
    Print #1, Format$(Now, "yyyy-mm-dd"), totalAmount
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("Format"));
        assert!(debug.contains("Now"));
        assert!(debug.contains("totalAmount"));
    }

    #[test]
    fn print_preserves_whitespace() {
        let source = r#"
Sub Test()
    Print   #1  ,   "Test"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn print_with_comment() {
        let source = r#"
Sub Test()
    Print #1, "Data" ' Write to file
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("LineComment"));
    }

    #[test]
    fn print_case_insensitive() {
        let source = r#"
Sub Test()
    PRINT #1, "Test"
    print #2, "test"
    PrInT #3, "TeSt"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        // Should have 3 print statements
        assert_eq!(debug.matches("PrintStatement").count(), 3);
    }

    #[test]
    fn print_with_line_continuation() {
        let source = r#"
Sub Test()
    Print #1, _
        "Line 1", _
        "Line 2"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("Underscore"));
    }

    #[test]
    fn print_numeric_expressions() {
        let source = r#"
Sub Test()
    Print #1, 42, 3.14, -100
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(
            debug.contains("IntegerLiteral")
                || debug.contains("SingleLiteral")
                || debug.contains("DoubleLiteral")
        );
    }

    #[test]
    fn print_boolean_values() {
        let source = r#"
Sub Test()
    Print #1, True, False
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("TrueKeyword"));
        assert!(debug.contains("FalseKeyword"));
    }

    #[test]
    fn print_with_spc_and_tab() {
        let source = r#"
Sub Test()
    Print #1, Spc(10); "Text"; Tab(20); "More"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("PrintStatement"));
        assert!(debug.contains("PrintKeyword"));
        assert!(debug.contains("Spc"));
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn print_multiple_statements() {
        let source = r#"
Sub Test()
    Print #1, "First"
    Print #2, "Second"
    Print #3, "Third"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(debug.matches("PrintStatement").count(), 3);
        assert!(debug.contains("PrintKeyword"));
    }
}
