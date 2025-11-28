//! VB6 `Tab` Function
//!
//! The `Tab` function is used in Print statements to position output at a specific column number.
//!
//! ## Syntax
//! ```vb6
//! Tab([column])
//! ```
//!
//! ## Parameters
//! - `column`: Optional. Numeric expression indicating the column number (1-based) at which to position the next character printed. If omitted, moves to the next print zone.
//!
//! ## Returns
//! Returns a special value used only in Print statements to control output position. It does not return a value for assignment or calculation.
//!
//! ## Remarks
//! - `Tab` is only meaningful within Print statements (e.g., `Print #1, Tab(10); "Hello"`).
//! - If `column` is omitted, output moves to the next print zone (every 14 columns by default).
//! - If `column` is less than the current print position, output moves to that column on the next line.
//! - If `column` is greater than the output line width, output starts at column 1 on the next line.
//! - `Tab` cannot be used in assignment or as a function value.
//! - `Tab` is not evaluated as a function in expressions outside Print context.
//! - `Tab` is not the same as the Tab key or character (Chr$(9)).
//! - In Print statements, `Tab` can be combined with `Spc` for advanced formatting.
//!
//! ## Typical Uses
//! 1. Aligning columns in printed output
//! 2. Formatting reports
//! 3. Creating tabular data in files
//! 4. Printing to the Immediate window
//! 5. Outputting to files with Print #
//! 6. Combining with `Spc` for custom spacing
//! 7. Printing headers and data in columns
//! 8. Generating formatted logs
//!
//! ## Basic Examples
//!
//! ### Example 1: Print with Tab
//! ```vb6
//! Print Tab(10); "Hello"
//! ```
//!
//! ### Example 2: Print to file with Tab
//! ```vb6
//! Print #1, Tab(20); "World"
//! ```
//!
//! ### Example 3: Print with omitted column
//! ```vb6
//! Print Tab; "Next zone"
//! ```
//!
//! ### Example 4: Print multiple columns
//! ```vb6
//! Print Tab(5); "A"; Tab(15); "B"; Tab(25); "C"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Print table header
//! ```vb6
//! Print Tab(1); "ID"; Tab(10); "Name"; Tab(30); "Score"
//! ```
//!
//! ### Pattern 2: Print data rows
//! ```vb6
//! For i = 1 To 10
//!     Print Tab(1); i; Tab(10); names(i); Tab(30); scores(i)
//! Next i
//! ```
//!
//! ### Pattern 3: Print with Spc
//! ```vb6
//! Print Tab(10); Spc(5); "Data"
//! ```
//!
//! ### Pattern 4: Print to Immediate window
//! ```vb6
//! Debug.Print Tab(15); "Debug info"
//! ```
//!
//! ### Pattern 5: Print to file
//! ```vb6
//! Print #1, Tab(8); "File data"
//! ```
//!
//! ### Pattern 6: Print with omitted column
//! ```vb6
//! Print Tab; "Default zone"
//! ```
//!
//! ### Pattern 7: Print with calculated column
//! ```vb6
//! Print Tab(i * 5); "Value"
//! ```
//!
//! ### Pattern 8: Print with variable
//! ```vb6
//! col = 12
//! Print Tab(col); "Text"
//! ```
//!
//! ### Pattern 9: Print with multiple Tab calls
//! ```vb6
//! Print Tab(5); "A"; Tab(15); "B"; Tab(25); "C"
//! ```
//!
//! ### Pattern 10: Print with Tab and Spc
//! ```vb6
//! Print Tab(10); Spc(3); "Mix"
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Print formatted report
//! ```vb6
//! Print Tab(1); "Header1"; Tab(20); "Header2"
//! For i = 1 To 5
//!     Print Tab(1); data1(i); Tab(20); data2(i)
//! Next i
//! ```
//!
//! ### Example 2: Print to file with dynamic columns
//! ```vb6
//! For i = 1 To 3
//!     Print #1, Tab(i * 10); "Col" & i
//! Next i
//! ```
//!
//! ### Example 3: Print with omitted column in loop
//! ```vb6
//! For i = 1 To 3
//!     Print Tab; "Row" & i
//! Next i
//! ```
//!
//! ### Example 4: Print with Tab and Spc for alignment
//! ```vb6
//! Print Tab(10); Spc(2); "Aligned"
//! ```
//!
//! ## Error Handling
//! - If `column` is less than 1, output starts at column 1 of the next line.
//! - If `column` is omitted, output moves to the next print zone.
//! - If `column` is greater than line width, output starts at column 1 of the next line.
//!
//! ## Performance Notes
//! - No performance impact; only affects output formatting.
//! - Used only in Print statements.
//!
//! ## Best Practices
//! 1. Use only in Print statements.
//! 2. Avoid using as a function in expressions.
//! 3. Use with Spc for custom spacing.
//! 4. Test output on different devices (screen, file).
//! 5. Use variables for dynamic columns.
//! 6. Document column positions for maintainability.
//! 7. Avoid negative or zero columns.
//! 8. Use for tabular data formatting.
//! 9. Combine with loops for tables.
//! 10. Use omitted column for default zones.
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Input | Returns |
//! |----------|---------|-------|---------|
//! | `Tab`    | Print position | column (optional) | Print formatting |
//! | `Spc`    | Print spaces | count | Print formatting |
//! | `Chr$(9)`| Tab character | n/a | String |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, VBScript
//! - Consistent across platforms
//! - Only for Print statements
//!
//! ## Limitations
//! - Not a function for assignment or calculation
//! - Only meaningful in Print context
//! - Not the same as the Tab character (Chr$(9))
//! - Cannot be used outside Print/Debug.Print/Print #

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_tab_basic() {
        let source = r#"
Sub Test()
    Print Tab(10); "Hello"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_print_to_file() {
        let source = r#"
Sub Test()
    Print #1, Tab(20); "World"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_omitted_column() {
        let source = r#"
Sub Test()
    Print Tab; "Next zone"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_multiple_columns() {
        let source = r#"
Sub Test()
    Print Tab(5); "A"; Tab(15); "B"; Tab(25); "C"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_table_header() {
        let source = r#"
Sub Test()
    Print Tab(1); "ID"; Tab(10); "Name"; Tab(30); "Score"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_data_rows() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Print Tab(1); i; Tab(10); names(i); Tab(30); scores(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_with_spc() {
        let source = r#"
Sub Test()
    Print Tab(10); Spc(5); "Data"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Tab(15); "Debug info"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_print_to_file_2() {
        let source = r#"
Sub Test()
    Print #1, Tab(8); "File data"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_omitted_column_2() {
        let source = r#"
Sub Test()
    Print Tab; "Default zone"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_calculated_column() {
        let source = r#"
Sub Test()
    Print Tab(i * 5); "Value"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_with_variable() {
        let source = r#"
Sub Test()
    col = 12
    Print Tab(col); "Text"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_multiple_tab_calls() {
        let source = r#"
Sub Test()
    Print Tab(5); "A"; Tab(15); "B"; Tab(25); "C"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_with_tab_and_spc() {
        let source = r#"
Sub Test()
    Print Tab(10); Spc(3); "Mix"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_formatted_report() {
        let source = r#"
Sub Test()
    Print Tab(1); "Header1"; Tab(20); "Header2"
    For i = 1 To 5
        Print Tab(1); data1(i); Tab(20); data2(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_dynamic_columns() {
        let source = r#"
Sub Test()
    For i = 1 To 3
        Print #1, Tab(i * 10); "Col" & i
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_omitted_column_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 3
        Print Tab; "Row" & i
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }

    #[test]
    fn test_tab_alignment() {
        let source = r#"
Sub Test()
    Print Tab(10); Spc(2); "Aligned"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tab"));
    }
}
