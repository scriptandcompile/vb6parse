//! # Width Statement
//!
//! Assigns an output line width to a file opened using the Open statement.
//!
//! ## Syntax
//!
//! ```vb
//! Width #filenumber, width
//! ```
//!
//! ## Parts
//!
//! - **filenumber**: Required. Any valid file number.
//! - **width**: Required. Numeric expression in the range 0–255, inclusive, that indicates how
//!   many characters appear on a line before a new line is started. If width equals 0, there is
//!   no limit to the length of a line. The default value for width is 0.
//!
//! ## Remarks
//!
//! - **Output Formatting**: The Width # statement is used with the Print # or Write # statements
//!   to control output formatting to files.
//! - **Line Length Control**: For files opened for sequential output, if the width of a line of
//!   output exceeds the value specified for width, a new line is automatically started.
//! - **No Effect on Input**: The Width # statement has no effect on files opened for input or
//!   binary access.
//! - **Zero Width**: Setting width to 0 means there is no line length limit, allowing continuous
//!   output without automatic line breaks.
//! - **Maximum Width**: The maximum width value is 255 characters.
//!
//! ## Examples
//!
//! ### Basic Width Setting
//!
//! ```vb
//! Open "output.txt" For Output As #1
//! Width #1, 80
//! Print #1, "This output will wrap at 80 characters"
//! Close #1
//! ```
//!
//! ### Set Unlimited Width
//!
//! ```vb
//! Open "data.csv" For Output As #2
//! Width #2, 0  ' No line length limit
//! Print #2, LongDataString
//! Close #2
//! ```
//!
//! ### Width with Multiple Files
//!
//! ```vb
//! Open "narrow.txt" For Output As #1
//! Open "wide.txt" For Output As #2
//! Width #1, 40
//! Width #2, 120
//! ```
//!
//! ### Dynamic Width Setting
//!
//! ```vb
//! Dim lineWidth As Integer
//! lineWidth = 80
//! Open "report.txt" For Output As #1
//! Width #1, lineWidth
//! ```
//!
//! ### Width for Formatted Output
//!
//! ```vb
//! Open "report.txt" For Output As #1
//! Width #1, 80
//! Print #1, Tab(10); "Header"
//! Print #1, Tab(10); String$(50, "-")
//! Close #1
//! ```
//!
//! ## Common Patterns
//!
//! ### Report Generation with Fixed Width
//!
//! ```vb
//! Sub GenerateReport()
//!     Open "report.txt" For Output As #1
//!     Width #1, 80
//!     
//!     Print #1, "Annual Sales Report"
//!     Print #1, String$(80, "=")
//!     ' ... report content ...
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### CSV Export (No Width Limit)
//!
//! ```vb
//! Sub ExportCSV()
//!     Open "export.csv" For Output As #1
//!     Width #1, 0  ' Allow unlimited line length
//!     
//!     For i = 1 To RecordCount
//!         Print #1, BuildCSVLine(i)
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Console-Style Output
//!
//! ```vb
//! Open "console.log" For Output As #1
//! Width #1, 80  ' Standard console width
//! Print #1, "System Log - "; Now()
//! Close #1
//! ```

use crate::parsers::SyntaxKind;

use super::super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Width # statement.
    ///
    /// The Width # statement assigns an output line width to a file opened using the Open statement.
    ///
    /// Syntax:
    /// ```vb
    /// Width #filenumber, width
    /// ```
    ///
    /// Example:
    /// ```vb
    /// Width #1, 80
    /// ```
    pub(in crate::parsers::cst) fn parse_width_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::WidthStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn width_simple() {
        let source = r#"
Sub Test()
    Width #1, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("WidthKeyword"));
    }

    #[test]
    fn width_at_module_level() {
        let source = r#"
Width #1, 80
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_with_zero() {
        let source = r#"
Sub Test()
    Width #1, 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("IntegerLiteral"));
    }

    #[test]
    fn width_with_variable() {
        let source = r#"
Sub Test()
    Width #1, lineWidth
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn width_with_expression() {
        let source = r#"
Sub Test()
    Width #1, maxWidth * 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_with_file_number_variable() {
        let source = r#"
Sub Test()
    Width #fileNum, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_max_value() {
        let source = r#"
Sub Test()
    Width #1, 255
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("IntegerLiteral"));
    }

    #[test]
    fn width_with_comment() {
        let source = r#"
Sub Test()
    Width #1, 80 ' Set standard console width
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn width_multiple_files() {
        let source = r#"
Sub Test()
    Width #1, 80
    Width #2, 120
    Width #3, 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WidthStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn width_with_spaces() {
        let source = r#"
Sub Test()
    Width  #1 ,  80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn width_in_if_statement() {
        let source = r#"
Sub Test()
    If openSuccess Then
        Width #1, 80
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn width_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Width #i, 80
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn width_after_open() {
        let source = r#"
Sub Test()
    Open "file.txt" For Output As #1
    Width #1, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("OpenStatement"));
    }

    #[test]
    fn width_before_print() {
        let source = r#"
Sub Test()
    Width #1, 80
    Print #1, "Output"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("PrintStatement"));
    }

    #[test]
    fn width_with_function_call() {
        let source = r#"
Sub Test()
    Width #1, GetLineWidth()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_with_constant() {
        let source = r#"
Sub Test()
    Width #1, MAX_WIDTH
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_in_with_block() {
        let source = r#"
Sub Test()
    With FileConfig
        Width #1, .LineWidth
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn width_sequential_calls() {
        let source = r#"
Sub Test()
    Open "file1.txt" For Output As #1
    Width #1, 80
    Open "file2.txt" For Output As #2
    Width #2, 120
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WidthStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn width_with_parenthesized_file_number() {
        let source = r#"
Sub Test()
    Width #(fileNum), 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_with_calculated_width() {
        let source = r#"
Sub Test()
    Width #1, screenWidth - marginLeft - marginRight
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_in_error_handler() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Width #1, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn width_with_type_suffix() {
        let source = r#"
Sub Test()
    Width #1, 80%
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_with_line_continuation() {
        let source = r#"
Sub Test()
    Width #1, _
        80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_case_insensitive() {
        let source = r#"
Sub Test()
    WIDTH #1, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
    }

    #[test]
    fn width_standard_values() {
        let source = r#"
Sub Test()
    Width #1, 40
    Width #2, 80
    Width #3, 132
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WidthStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn width_with_file_freefile() {
        let source = r#"
Sub Test()
    Dim fn As Integer
    fn = FreeFile
    Width #fn, 80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn width_in_select_case() {
        let source = r#"
Sub Test()
    Select Case outputType
        Case 1
            Width #1, 80
        Case 2
            Width #1, 132
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WidthStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn width_preserves_formatting() {
        let source = r#"
Sub Test()
    Width    #1   ,    80
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn width_in_nested_control_structures() {
        let source = r#"
Sub Test()
    If fileOpen Then
        For i = 1 To 10
            Width #i, 80
        Next i
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WidthStatement"));
        assert!(debug.contains("IfStatement"));
        assert!(debug.contains("ForStatement"));
    }
}
