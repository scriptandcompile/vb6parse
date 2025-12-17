//! # Write Statement
//!
//! Writes data to a sequential file.
//!
//! ## Syntax
//!
//! ```vb
//! Write #filenumber, [outputlist]
//! ```
//!
//! ## Parts
//!
//! - **filenumber**: Required. Any valid file number.
//! - **outputlist**: Optional. One or more comma-delimited numeric expressions or string expressions
//!   to write to a file.
//!
//! ## Remarks
//!
//! - **Data Formatting**: Data written with Write # is usually read from a file with Input #.
//! - **Delimiters**: The Write # statement inserts commas between items and quotation marks around
//!   strings as they are written to the file. You don't have to put explicit delimiters in the list.
//! - **Universal Data**: Write # writes data in a universal format that can be read by Input # regardless
//!   of the locale settings.
//! - **Numeric Data**: Numeric data is written with a period (.) as the decimal separator.
//! - **Boolean Values**: Boolean data is written as #TRUE# or #FALSE#.
//! - **Date Values**: Date data is written using the universal date format: #yyyy-mm-dd hh:mm:ss#
//! - **Empty Values**: If outputlist data is Empty, nothing is written. However, if outputlist data is
//!   Null, #NULL# is written.
//! - **Error Data**: Error values are written as #ERROR errorcode#. The number sign (#) ensures the keyword
//!   is not confused with a variable name.
//! - **Comparison with Print #**: Unlike Print #, Write # inserts commas between items and quotes around
//!   strings automatically.
//!
//! ## Examples
//!
//! ### Write Simple Data
//!
//! ```vb
//! Open "test.txt" For Output As #1
//! Write #1, "Hello", 42, True
//! Close #1
//! ' File contents: "Hello",42,#TRUE#
//! ```
//!
//! ### Write Multiple Lines
//!
//! ```vb
//! Open "data.txt" For Output As #1
//! For i = 1 To 10
//!     Write #1, i, i * i, i * i * i
//! Next i
//! Close #1
//! ```
//!
//! ### Write Mixed Data Types
//!
//! ```vb
//! Open "record.txt" For Output As #1
//! Write #1, "John Doe", 30, #1/1/1995#, True
//! Close #1
//! ```
//!
//! ### Write Without Data (New Line)
//!
//! ```vb
//! Open "output.txt" For Output As #1
//! Write #1, "First line"
//! Write #1
//! Write #1, "Third line"
//! Close #1
//! ```
//!
//! ### Write Null and Empty
//!
//! ```vb
//! Open "test.txt" For Output As #1
//! Write #1, Null, Empty, "data"
//! Close #1
//! ' File contents: #NULL#,,"data"
//! ```
//!
//! ### Write Error Values
//!
//! ```vb
//! Open "errors.txt" For Output As #1
//! Write #1, CVErr(2007)
//! Close #1
//! ' File contents: #ERROR 2007#
//! ```
//!
//! ## Common Patterns
//!
//! ### Export Data to CSV-like Format
//!
//! ```vb
//! Sub ExportData()
//!     Open "export.txt" For Output As #1
//!     
//!     ' Write header
//!     Write #1, "Name", "Age", "City"
//!     
//!     ' Write data rows
//!     For i = 0 To UBound(employees)
//!         Write #1, employees(i).Name, employees(i).Age, employees(i).City
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Database Records
//!
//! ```vb
//! Sub SaveRecords()
//!     Open "records.dat" For Output As #1
//!     
//!     Do Until rs.EOF
//!         Write #1, rs!ID, rs!Name, rs!Date, rs!Active
//!         rs.MoveNext
//!     Loop
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Configuration Data
//!
//! ```vb
//! Sub SaveConfig()
//!     Open "config.dat" For Output As #1
//!     Write #1, appName, version, lastRun, isRegistered
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Array Data
//!
//! ```vb
//! Sub WriteArray()
//!     Open "array.dat" For Output As #1
//!     
//!     For i = LBound(data) To UBound(data)
//!         Write #1, data(i)
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Append Data to Existing File
//!
//! ```vb
//! Sub AppendRecord()
//!     Open "log.txt" For Append As #1
//!     Write #1, Now(), userName, action, details
//!     Close #1
//! End Sub
//! ```

use crate::parsers::SyntaxKind;

use super::super::Parser;

impl Parser<'_> {
    /// Parse a Write # statement.
    ///
    /// The Write # statement writes data to a sequential file with automatic
    /// formatting: commas between items and quotation marks around strings.
    ///
    /// Syntax:
    /// ```vb
    /// Write #filenumber, [outputlist]
    /// ```
    ///
    /// Example:
    /// ```vb
    /// Write #1, "Hello", 42, True
    /// ```
    pub(in crate::parsers::cst) fn parse_write_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::WriteStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn write_simple() {
        let source = r#"
Sub Test()
    Write #1, "Hello"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("WriteKeyword"));
    }

    #[test]
    fn write_at_module_level() {
        let source = r#"
Write #1, "data"
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_multiple_values() {
        let source = r#"
Sub Test()
    Write #1, "Name", 42, True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("Name"));
    }

    #[test]
    fn write_no_data() {
        let source = r#"
Sub Test()
    Write #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_with_variables() {
        let source = r#"
Sub Test()
    Write #1, name, age, city
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("name"));
        assert!(debug.contains("age"));
        assert!(debug.contains("city"));
    }

    #[test]
    fn write_with_expressions() {
        let source = r#"
Sub Test()
    Write #1, x + y, total * 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_with_file_number_variable() {
        let source = r#"
Sub Test()
    Write #fileNum, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("fileNum"));
    }

    #[test]
    fn write_with_comment() {
        let source = r#"
Sub Test()
    Write #1, data ' Write data to file
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn write_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Write #1, i, i * i
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn write_with_string_literal() {
        let source = r#"
Sub Test()
    Write #1, "Hello, World!", "Data"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("Hello"));
    }

    #[test]
    fn write_with_numeric_literals() {
        let source = r#"
Sub Test()
    Write #1, 42, 3.14, -100
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_with_boolean() {
        let source = r#"
Sub Test()
    Write #1, True, False
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("True"));
        assert!(debug.contains("False"));
    }

    #[test]
    fn write_with_date() {
        let source = r#"
Sub Test()
    Write #1, #1/1/2025#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_with_null() {
        let source = r#"
Sub Test()
    Write #1, Null, Empty
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_with_object_property() {
        let source = r#"
Sub Test()
    Write #1, obj.Name, obj.Value
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("obj"));
    }

    #[test]
    fn write_with_array_access() {
        let source = r#"
Sub Test()
    Write #1, arr(i), arr(j)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("arr"));
    }

    #[test]
    fn write_with_function_call() {
        let source = r#"
Sub Test()
    Write #1, GetValue(), ProcessData()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("GetValue"));
    }

    #[test]
    fn write_multiple_statements() {
        let source = r#"
Sub Test()
    Write #1, "Line 1"
    Write #1, "Line 2"
    Write #1, "Line 3"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WriteStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn write_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        Write #1, data
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn write_in_do_loop() {
        let source = r#"
Sub Test()
    Do Until EOF(1)
        Write #2, currentRecord
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn write_with_recordset() {
        let source = r#"
Sub Test()
    Write #1, rs!Name, rs!Age
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_preserves_whitespace() {
        let source = r#"
Sub Test()
    Write  #1 ,  data1 ,  data2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn write_with_line_continuation() {
        let source = r#"
Sub Test()
    Write #1, _
        field1, _
        field2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_in_select_case() {
        let source = r#"
Sub Test()
    Select Case recordType
        Case 1
            Write #1, "Type A", data
        Case 2
            Write #1, "Type B", data
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WriteStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn write_with_now_function() {
        let source = r#"
Sub Test()
    Write #1, Now(), data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("Now"));
    }

    #[test]
    fn write_in_with_block() {
        let source = r#"
Sub Test()
    With record
        Write #1, .Name, .Value
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn write_case_insensitive() {
        let source = r#"
Sub Test()
    WRITE #1, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }

    #[test]
    fn write_in_error_handler() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Write #1, errorData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn write_with_freefile() {
        let source = r#"
Sub Test()
    Dim fn As Integer
    fn = FreeFile
    Write #fn, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn write_sequential_values() {
        let source = r#"
Sub Test()
    For i = 1 To 100
        Write #1, i, i * 2, i ^ 2
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WriteStatement"));
    }
}
