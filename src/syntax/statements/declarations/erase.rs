//! Erase statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Erase statements:
//! - `Erase` - Reinitialize fixed-size array elements and deallocate dynamic arrays
//!
//! # Erase Statement
//!
//! The Erase statement is used to reinitialize the elements of fixed-size arrays
//! and to release storage space used by dynamic arrays.
//!
//! ## Syntax
//! ```vb
//! Erase arraylist
//! ```
//!
//! ## Behavior
//! - For fixed-size arrays: Reinitializes the elements to their default values
//!   (0 for numeric types, "" for strings, Nothing for objects)
//! - For dynamic arrays: Deallocates the memory used by the array
//!
//! ## Examples
//! ```vb
//! Erase myArray
//! Erase array1, array2, array3
//! ```
//!
//! ## Remarks
//! - The arraylist argument is a list of one or more comma-delimited array variable names
//! - After erasing a dynamic array, you must use `ReDim` to reallocate it before using again
//! - Erasing a fixed-size array does not deallocate memory, just resets values
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/erase-statement)

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse an Erase statement: Erase array1 [, array2] ...
    ///
    /// VB6 Erase statement syntax:
    /// - Erase arraylist
    ///
    /// The Erase statement is used to reinitialize the elements of fixed-size arrays
    /// and to release storage space used by dynamic arrays.
    ///
    /// The arraylist argument is a list of one or more comma-delimited array variable names.
    ///
    /// Behavior:
    /// - For fixed-size arrays: Reinitializes the elements to their default values
    ///   (0 for numeric types, "" for strings, Nothing for objects)
    /// - For dynamic arrays: Deallocates the memory used by the array
    ///
    /// Examples:
    /// ```vb
    /// Erase myArray
    /// Erase array1, array2, array3
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/erase-statement)
    pub(crate) fn parse_erase_statement(&mut self) {
        // if we are now parsing an erase statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::EraseStatement.to_raw());

        // Consume "Erase" keyword
        self.consume_token();

        // Consume everything until newline (array names, commas, etc.)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // EraseStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn erase_simple_array() {
        let source = r"
Sub Test()
    Erase myArray
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_multiple_arrays() {
        let source = r"
Sub Test()
    Erase array1, array2, array3
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_at_module_level() {
        let source = "Erase globalArray\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_preserves_whitespace() {
        let source = "    Erase    myArray    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_with_comment() {
        let source = r"
Sub Test()
    Erase tempArray ' Free up memory
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_in_if_statement() {
        let source = r"
Sub Cleanup()
    If shouldClear Then
        Erase dataArray
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_inline_if() {
        let source = r"
Sub Test()
    If resetFlag Then Erase buffer
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Erase tempArrays(i)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_with_parentheses() {
        let source = r"
Sub Test()
    Erase myArray()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_erase_statements() {
        let source = r"
Sub Test()
    Erase array1
    DoSomething
    Erase array2
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Erase dynamicArray
    If Err.Number <> 0 Then
        MsgBox "Error erasing array"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_after_redim() {
        let source = r"
Sub Test()
    ReDim myArray(100)
    ' Use the array
    Erase myArray
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn erase_complex_array_list() {
        let source = r"
Sub Test()
    Erase buffer1, buffer2, cache(), tempData
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/erase");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
